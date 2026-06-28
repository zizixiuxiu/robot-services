from datetime import datetime, date
from sqlalchemy import String, DateTime, Date, Text, ForeignKey, func
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db import Base


class InboundOrder(Base):
    __tablename__ = "inbound_orders"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    inbound_no: Mapped[str] = mapped_column(String(255), nullable=False, unique=True, index=True)
    inbound_date: Mapped[date] = mapped_column(Date, nullable=False)
    related_order_id: Mapped[int | None] = mapped_column(ForeignKey("orders.id"), nullable=True)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )

    related_order: Mapped["Order | None"] = relationship("Order")
    lines: Mapped[list["InboundOrderLine"]] = relationship(
        "InboundOrderLine", back_populates="inbound_order", cascade="all, delete-orphan", lazy="selectin"
    )
