from datetime import date, datetime
from sqlalchemy.orm import Session
from app.models.inbound_order import InboundOrder
from app.models.outbound_order import OutboundOrder
from app.models.inbound_order_line import InboundOrderLine
from app.models.order import Order
from app.models.order_line import OrderLine
from app.schemas.inbound import InboundOrderCreate
from app.services.orders import ORDER_STATUS_PRODUCING, ORDER_STATUS_PARTIALLY_INBOUND, ORDER_STATUS_IN_STOCK, ORDER_STATUS_PARTIALLY_SHIPPED
from app.services.orders import OrderService
from app.services.stock import StockService


class InboundService:
    @staticmethod
    def generate_inbound_no(db: Session) -> str:
        prefix = f"INB-{datetime.now().strftime('%Y%m%d')}"
        latest_no = (
            db.query(InboundOrder)
            .filter(InboundOrder.inbound_no.like(f"{prefix}-%"))
            .order_by(InboundOrder.inbound_no.desc())
            .with_entities(InboundOrder.inbound_no)
            .first()
        )
        seq = int(latest_no[0].rsplit("-", 1)[-1]) + 1 if latest_no else 1
        return f"{prefix}-{seq:04d}"

    @staticmethod
    def list_inbounds(db: Session):
        return (
            db.query(InboundOrder)
            .order_by(InboundOrder.id.desc())
            .all()
        )

    @staticmethod
    def list_inbounds_for_outbound(db: Session):
        inbounds = (
            db.query(InboundOrder)
            .order_by(InboundOrder.id.desc())
            .all()
        )
        result = []
        for inbound in inbounds:
            if not inbound.related_order:
                result.append(inbound)
                continue
            if inbound.related_order.status in [ORDER_STATUS_PARTIALLY_INBOUND, ORDER_STATUS_IN_STOCK, ORDER_STATUS_PARTIALLY_SHIPPED]:
                result.append(inbound)
        return result

    @staticmethod
    def get_inbound(db: Session, inbound_id: int):
        return db.query(InboundOrder).filter(InboundOrder.id == inbound_id).first()

    @staticmethod
    def create_inbound(db: Session, data: InboundOrderCreate):
        related_order = None
        order_lines = []
        if data.related_order_id:
            related_order = db.query(Order).filter(Order.id == data.related_order_id).first()
            if not related_order:
                raise ValueError("Linked order does not exist.")
            if related_order.status not in [ORDER_STATUS_PRODUCING, ORDER_STATUS_PARTIALLY_INBOUND]:
                raise ValueError("Only producing or partially inbounded orders can continue inbound.")

            order_lines = db.query(OrderLine).filter(OrderLine.order_id == related_order.id).all()
            line_map = {line.product_id: line for line in order_lines}

            inbound_qty_by_product: dict[int, int] = {}
            for line_data in data.lines:
                inbound_qty_by_product[line_data.product_id] = inbound_qty_by_product.get(line_data.product_id, 0) + line_data.qty

            if not inbound_qty_by_product:
                raise ValueError("Linked inbound requires at least one valid line.")

            for product_id, inbound_qty in inbound_qty_by_product.items():
                order_line = line_map.get(product_id)
                if not order_line:
                    raise ValueError("Inbound product must belong to the linked order.")
                remaining_inbound_qty = order_line.qty - (order_line.inbound_qty or 0)
                if inbound_qty > remaining_inbound_qty:
                    raise ValueError("Inbound quantity exceeds the remaining quantity of the linked order.")

            for order_line in order_lines:
                if inbound_qty_by_product.get(order_line.product_id, 0) <= 0:
                    continue

        order = InboundOrder(
            inbound_no=data.inbound_no,
            inbound_date=data.inbound_date,
            related_order_id=data.related_order_id,
            remark=data.remark,
        )
        db.add(order)
        db.commit()
        db.refresh(order)

        for line_data in data.lines:
            line = InboundOrderLine(
                inbound_order_id=order.id,
                product_id=line_data.product_id,
                qty=line_data.qty,
                remark=line_data.remark,
            )
            db.add(line)

            StockService.stock_in_with_biz(
                db=db,
                product_id=line_data.product_id,
                qty=line_data.qty,
                remark=line_data.remark,
                biz_type="inbound",
                biz_id=order.id,
            )

        if related_order:
            line_map = {line.product_id: line for line in order_lines}
            for line_data in data.lines:
                linked_line = line_map.get(line_data.product_id)
                if linked_line:
                    linked_line.inbound_qty = (linked_line.inbound_qty or 0) + line_data.qty

            all_inbounded = all((line.inbound_qty or 0) >= line.qty for line in order_lines)
            any_inbounded = any((line.inbound_qty or 0) > 0 for line in order_lines)
            any_shipped = any((line.shipped_qty or 0) > 0 for line in order_lines)

            if all_inbounded and any_shipped:
                related_order.status = ORDER_STATUS_PARTIALLY_SHIPPED
            elif all_inbounded:
                related_order.status = ORDER_STATUS_IN_STOCK
            elif any_inbounded:
                related_order.status = ORDER_STATUS_PARTIALLY_INBOUND

        db.commit()
        db.refresh(order)
        return order

    @staticmethod
    def delete_inbound(db: Session, inbound: InboundOrder):
        outbound_count = db.query(OutboundOrder).filter(OutboundOrder.related_inbound_id == inbound.id).count()
        if outbound_count:
            raise ValueError("该入库单已关联出库单，不能直接删除。")

        qty_by_product: dict[int, int] = {}
        for line in inbound.lines:
            qty_by_product[line.product_id] = qty_by_product.get(line.product_id, 0) + line.qty

        for product_id, qty in qty_by_product.items():
            balance = StockService.get_balance_by_product(db, product_id)
            if not balance or balance.qty < qty:
                raise ValueError("删除该入库单会导致库存为负，不能删除。")

        related_order = inbound.related_order
        if related_order:
            order_line_map = {line.product_id: line for line in related_order.lines}
            for product_id, qty in qty_by_product.items():
                order_line = order_line_map.get(product_id)
                if not order_line:
                    continue
                new_inbound_qty = max((order_line.inbound_qty or 0) - qty, 0)
                if (order_line.shipped_qty or 0) > new_inbound_qty:
                    raise ValueError("该订单已有出库数量，删除入库单会导致进度不一致。")

        for line in inbound.lines:
            StockService.stock_out_with_biz(
                db=db,
                product_id=line.product_id,
                qty=line.qty,
                remark=f"删除入库单 {inbound.inbound_no}",
                biz_type="delete_inbound",
                biz_id=inbound.id,
            )

        if related_order:
            order_line_map = {line.product_id: line for line in related_order.lines}
            for product_id, qty in qty_by_product.items():
                order_line = order_line_map.get(product_id)
                if order_line:
                    order_line.inbound_qty = max((order_line.inbound_qty or 0) - qty, 0)
            OrderService.refresh_order_status_from_lines(related_order)

        db.delete(inbound)
        db.commit()
