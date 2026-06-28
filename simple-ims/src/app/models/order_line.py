from decimal import Decimal
from sqlalchemy import Integer, ForeignKey, Numeric, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db import Base


class OrderLine(Base):
    __tablename__ = "order_lines"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"), nullable=False)
    product_id: Mapped[int] = mapped_column(ForeignKey("products.id"), nullable=False)
    qty: Mapped[int] = mapped_column(Integer, nullable=False)
    unit_price: Mapped[Decimal] = mapped_column(Numeric(12, 2), nullable=False, default=0)
    total_price: Mapped[Decimal] = mapped_column(Numeric(12, 2), nullable=False, default=0)
    substrate_category: Mapped[str] = mapped_column(Text, nullable=False, default="")
    surface_type: Mapped[str] = mapped_column(Text, nullable=False, default="")
    inbound_qty: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    shipped_qty: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)

    order: Mapped["Order"] = relationship("Order", back_populates="lines")
    product: Mapped["Product"] = relationship("Product")
