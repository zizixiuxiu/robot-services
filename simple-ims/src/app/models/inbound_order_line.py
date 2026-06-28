from sqlalchemy import Integer, ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db import Base


class InboundOrderLine(Base):
    __tablename__ = "inbound_order_lines"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    inbound_order_id: Mapped[int] = mapped_column(ForeignKey("inbound_orders.id"), nullable=False)
    product_id: Mapped[int] = mapped_column(ForeignKey("products.id"), nullable=False)
    qty: Mapped[int] = mapped_column(Integer, nullable=False)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)

    inbound_order: Mapped["InboundOrder"] = relationship("InboundOrder", back_populates="lines")
    product: Mapped["Product"] = relationship("Product")
