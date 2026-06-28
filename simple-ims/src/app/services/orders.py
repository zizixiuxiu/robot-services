from datetime import date, datetime
from decimal import Decimal
import random
from sqlalchemy.orm import Session
from app.models.inbound_order import InboundOrder
from app.models.order import Order
from app.models.order_line import OrderLine
from app.models.outbound_order import OutboundOrder
from app.schemas.order import OrderCreate, OrderUpdate

ORDER_STATUS_PRODUCING = "producing"
ORDER_STATUS_PARTIALLY_INBOUND = "partially_inbound"
ORDER_STATUS_IN_STOCK = "in_stock"
ORDER_STATUS_PARTIALLY_SHIPPED = "partially_shipped"
ORDER_STATUS_SHIPPED = "shipped"
ORDER_STATUS_CANCELLED = "cancelled"


class OrderService:
    @staticmethod
    def refresh_order_status_from_lines(order: Order) -> None:
        lines = order.lines or []
        if not lines:
            order.status = ORDER_STATUS_PRODUCING
            return

        all_shipped = all((line.shipped_qty or 0) >= line.qty for line in lines)
        any_shipped = any((line.shipped_qty or 0) > 0 for line in lines)
        all_inbounded = all((line.inbound_qty or 0) >= line.qty for line in lines)
        any_inbounded = any((line.inbound_qty or 0) > 0 for line in lines)

        if all_shipped:
            order.status = ORDER_STATUS_SHIPPED
        elif any_shipped:
            order.status = ORDER_STATUS_PARTIALLY_SHIPPED
        elif all_inbounded:
            order.status = ORDER_STATUS_IN_STOCK
        elif any_inbounded:
            order.status = ORDER_STATUS_PARTIALLY_INBOUND
        else:
            order.status = ORDER_STATUS_PRODUCING

    @staticmethod
    def calculate_line_total(qty: int, unit_price: Decimal) -> Decimal:
        return (Decimal(qty) * unit_price).quantize(Decimal("0.01"))

    @staticmethod
    def generate_order_no(db: Session) -> str:
        prefix = f"ORD-{datetime.now().strftime('%Y%m%d')}"
        latest_no = (
            db.query(Order)
            .filter(Order.order_no.like(f"{prefix}-%"))
            .order_by(Order.order_no.desc())
            .with_entities(Order.order_no)
            .first()
        )
        seq = int(latest_no[0].rsplit("-", 1)[-1]) + 1 if latest_no else 1
        return f"{prefix}-{seq:04d}"

    @staticmethod
    def list_orders(db: Session):
        return (
            db.query(Order)
            .order_by(Order.id.desc())
            .all()
        )

    @staticmethod
    def list_orders_for_inbound(db: Session):
        return (
            db.query(Order)
            .filter(Order.status == ORDER_STATUS_PRODUCING)
            .union(
                db.query(Order).filter(Order.status == ORDER_STATUS_PARTIALLY_INBOUND)
            )
            .order_by(Order.id.desc())
            .all()
        )

    @staticmethod
    def list_orders_for_outbound(db: Session):
        return (
            db.query(Order)
            .filter(Order.status.in_([ORDER_STATUS_PARTIALLY_INBOUND, ORDER_STATUS_IN_STOCK, ORDER_STATUS_PARTIALLY_SHIPPED]))
            .order_by(Order.id.desc())
            .all()
        )

    @staticmethod
    def get_order(db: Session, order_id: int):
        return db.query(Order).filter(Order.id == order_id).first()

    @staticmethod
    def create_order(db: Session, data: OrderCreate):
        order = Order(
            order_no=data.order_no,
            order_date=data.order_date,
            due_date=data.due_date,
            owner_user_id=data.owner_user_id,
            status=ORDER_STATUS_PRODUCING,
            remark=data.remark,
        )
        db.add(order)
        db.commit()
        db.refresh(order)

        for line_data in data.lines:
            total_price = OrderService.calculate_line_total(line_data.qty, line_data.unit_price)
            line = OrderLine(
                order_id=order.id,
                product_id=line_data.product_id,
                qty=line_data.qty,
                unit_price=line_data.unit_price,
                total_price=total_price,
                substrate_category=line_data.substrate_category,
                surface_type=line_data.surface_type,
                inbound_qty=0,
                shipped_qty=0,
                remark=line_data.remark,
            )
            db.add(line)

        db.commit()
        db.refresh(order)
        return order

    @staticmethod
    def delete_order(db: Session, order: Order):
        inbound_count = db.query(InboundOrder).filter(InboundOrder.related_order_id == order.id).count()
        outbound_count = db.query(OutboundOrder).filter(OutboundOrder.related_order_id == order.id).count()
        if inbound_count or outbound_count:
            raise ValueError("该订单已关联入库单或出库单，不能直接删除。")
        db.delete(order)
        db.commit()

    @staticmethod
    def update_order(db: Session, order: Order, data: OrderUpdate):
        if data.order_date is not None:
            order.order_date = data.order_date
        if data.due_date is not None:
            order.due_date = data.due_date
        if data.owner_user_id is not None:
            order.owner_user_id = data.owner_user_id
        if data.remark is not None:
            order.remark = data.remark

        if data.lines is not None:
            # remove existing lines and recreate
            for line in order.lines:
                db.delete(line)
            for line_data in data.lines:
                total_price = OrderService.calculate_line_total(line_data.qty, line_data.unit_price)
                line = OrderLine(
                    order_id=order.id,
                    product_id=line_data.product_id,
                    qty=line_data.qty,
                    unit_price=line_data.unit_price,
                    total_price=total_price,
                    substrate_category=line_data.substrate_category,
                    surface_type=line_data.surface_type,
                    inbound_qty=0,
                    shipped_qty=0,
                    remark=line_data.remark,
                )
                db.add(line)

        db.commit()
        db.refresh(order)
        return order
