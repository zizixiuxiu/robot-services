from datetime import date, datetime
from decimal import Decimal
from sqlalchemy.orm import Session
from app.models.outbound_order import OutboundOrder
from app.models.outbound_order_line import OutboundOrderLine
from app.models.order import Order
from app.models.order_line import OrderLine
from app.models.inbound_order import InboundOrder
from app.schemas.outbound import OutboundOrderCreate
from app.services.orders import ORDER_STATUS_PARTIALLY_INBOUND, ORDER_STATUS_IN_STOCK, ORDER_STATUS_PARTIALLY_SHIPPED, ORDER_STATUS_SHIPPED
from app.services.orders import OrderService
from app.services.stock import StockService


class OutboundService:
    @staticmethod
    def calculate_line_total(qty: int, unit_price: Decimal) -> Decimal:
        return (Decimal(qty) * unit_price).quantize(Decimal("0.01"))

    @staticmethod
    def generate_outbound_no(db: Session) -> str:
        prefix = f"OUT-{datetime.now().strftime('%Y%m%d')}"
        latest_no = (
            db.query(OutboundOrder)
            .filter(OutboundOrder.outbound_no.like(f"{prefix}-%"))
            .order_by(OutboundOrder.outbound_no.desc())
            .with_entities(OutboundOrder.outbound_no)
            .first()
        )
        seq = int(latest_no[0].rsplit("-", 1)[-1]) + 1 if latest_no else 1
        return f"{prefix}-{seq:04d}"

    @staticmethod
    def list_outbounds(db: Session):
        return (
            db.query(OutboundOrder)
            .order_by(OutboundOrder.id.desc())
            .all()
        )

    @staticmethod
    def get_outbound(db: Session, outbound_id: int):
        return db.query(OutboundOrder).filter(OutboundOrder.id == outbound_id).first()

    @staticmethod
    def create_outbound(db: Session, data: OutboundOrderCreate):
        linked_order = None
        if data.related_inbound_id:
            inbound = db.query(InboundOrder).filter(InboundOrder.id == data.related_inbound_id).first()
            if not inbound:
                raise ValueError("Linked inbound does not exist.")
            if inbound.related_order_id:
                linked_order = db.query(Order).filter(Order.id == inbound.related_order_id).first()
                data.related_order_id = inbound.related_order_id
        elif data.related_order_id:
            linked_order = db.query(Order).filter(Order.id == data.related_order_id).first()

        if linked_order:
            if linked_order.status not in [ORDER_STATUS_PARTIALLY_INBOUND, ORDER_STATUS_IN_STOCK, ORDER_STATUS_PARTIALLY_SHIPPED]:
                raise ValueError("Only inbounded orders can be shipped.")
            if linked_order.owner_user_id != data.owner_user_id:
                raise ValueError("Outbound owner must match the linked order owner.")

            order_lines = db.query(OrderLine).filter(OrderLine.order_id == linked_order.id).all()
            line_map = {line.product_id: line for line in order_lines}
            for line_data in data.lines:
                order_line = line_map.get(line_data.product_id)
                if not order_line:
                    raise ValueError("Outbound product must belong to the linked order.")
                remaining_qty = (order_line.inbound_qty or 0) - (order_line.shipped_qty or 0)
                if line_data.qty > remaining_qty:
                    raise ValueError("Outbound quantity exceeds the remaining quantity of the linked order.")
                line_data.unit_price = order_line.unit_price or Decimal("0")
                line_data.total_price = OutboundService.calculate_line_total(line_data.qty, line_data.unit_price)

        # Pre-validate stock availability
        for line_data in data.lines:
            balance = StockService.get_balance_by_product(db, line_data.product_id)
            before_qty = balance.qty if balance else 0
            if before_qty < line_data.qty:
                raise ValueError(f"Insufficient stock for product {line_data.product_id}")

        order = OutboundOrder(
            outbound_no=data.outbound_no,
            outbound_date=data.outbound_date,
            owner_user_id=data.owner_user_id,
            related_order_id=data.related_order_id,
            related_inbound_id=data.related_inbound_id,
            remark=data.remark,
        )
        db.add(order)
        db.commit()
        db.refresh(order)

        for line_data in data.lines:
            total_price = OutboundService.calculate_line_total(line_data.qty, line_data.unit_price)
            line = OutboundOrderLine(
                outbound_order_id=order.id,
                product_id=line_data.product_id,
                qty=line_data.qty,
                unit_price=line_data.unit_price,
                total_price=total_price,
                remark=line_data.remark,
            )
            db.add(line)

            StockService.stock_out_with_biz(
                db=db,
                product_id=line_data.product_id,
                qty=line_data.qty,
                remark=line_data.remark,
                biz_type="outbound",
                biz_id=order.id,
            )

        # Update linked order shipped quantities and status
        if linked_order:
            order_lines = db.query(OrderLine).filter(OrderLine.order_id == linked_order.id).all()
            line_map = {ol.product_id: ol for ol in order_lines}

            for line_data in data.lines:
                ol = line_map.get(line_data.product_id)
                if ol:
                    ol.shipped_qty = (ol.shipped_qty or 0) + line_data.qty

            all_shipped = all(ol.shipped_qty >= ol.qty for ol in order_lines)
            any_shipped = any(ol.shipped_qty > 0 for ol in order_lines)

            if all_shipped:
                linked_order.status = ORDER_STATUS_SHIPPED
            elif any_shipped:
                linked_order.status = ORDER_STATUS_PARTIALLY_SHIPPED

        db.commit()
        db.refresh(order)
        return order

    @staticmethod
    def delete_outbound(db: Session, outbound: OutboundOrder):
        linked_order = outbound.related_order
        qty_by_product: dict[int, int] = {}
        for line in outbound.lines:
            qty_by_product[line.product_id] = qty_by_product.get(line.product_id, 0) + line.qty

        for line in outbound.lines:
            StockService.stock_in_with_biz(
                db=db,
                product_id=line.product_id,
                qty=line.qty,
                remark=f"删除出库单 {outbound.outbound_no}",
                biz_type="delete_outbound",
                biz_id=outbound.id,
            )

        if linked_order:
            order_line_map = {line.product_id: line for line in linked_order.lines}
            for product_id, qty in qty_by_product.items():
                order_line = order_line_map.get(product_id)
                if order_line:
                    order_line.shipped_qty = max((order_line.shipped_qty or 0) - qty, 0)
            OrderService.refresh_order_status_from_lines(linked_order)

        db.delete(outbound)
        db.commit()
