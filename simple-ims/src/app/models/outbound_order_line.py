from decimal import Decimal
from sqlalchemy import Integer, ForeignKey, Numeric, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db import Base


class OutboundOrderLine(Base):
    __tablename__ = "outbound_order_lines"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    outbound_order_id: Mapped[int] = mapped_column(ForeignKey("outbound_orders.id"), nullable=False)
    product_id: Mapped[int] = mapped_column(ForeignKey("products.id"), nullable=False)
    qty: Mapped[int] = mapped_column(Integer, nullable=False)
    unit_price: Mapped[Decimal] = mapped_column(Numeric(12, 2), nullable=False, default=0)
    total_price: Mapped[Decimal] = mapped_column(Numeric(12, 2), nullable=False, default=0)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)

    outbound_order: Mapped["OutboundOrder"] = relationship("OutboundOrder", back_populates="lines")
    product: Mapped["Product"] = relationship("Product")
