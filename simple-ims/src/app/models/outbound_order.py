from datetime import datetime, date
from sqlalchemy import String, DateTime, Date, Text, ForeignKey, func
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db import Base


class OutboundOrder(Base):
    __tablename__ = "outbound_orders"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    outbound_no: Mapped[str] = mapped_column(String(255), nullable=False, unique=True, index=True)
    outbound_date: Mapped[date] = mapped_column(Date, nullable=False)
    owner_user_id: Mapped[int] = mapped_column(ForeignKey("owners.id"), nullable=False)
    related_order_id: Mapped[int | None] = mapped_column(ForeignKey("orders.id"), nullable=True)
    related_inbound_id: Mapped[int | None] = mapped_column(ForeignKey("inbound_orders.id"), nullable=True)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )

    owner: Mapped["Owner"] = relationship("Owner")
    related_order: Mapped["Order | None"] = relationship("Order")
    related_inbound: Mapped["InboundOrder | None"] = relationship("InboundOrder")
    lines: Mapped[list["OutboundOrderLine"]] = relationship(
        "OutboundOrderLine", back_populates="outbound_order", cascade="all, delete-orphan", lazy="selectin"
    )
