from datetime import datetime
from sqlalchemy import ForeignKey, String, DateTime, func, Text
from sqlalchemy.orm import Mapped, mapped_column
from app.db import Base


class StockMovement(Base):
    __tablename__ = "stock_movements"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    product_id: Mapped[int] = mapped_column(
        ForeignKey("products.id"), nullable=False, index=True
    )
    change_qty: Mapped[int] = mapped_column(nullable=False)
    before_qty: Mapped[int] = mapped_column(nullable=False)
    after_qty: Mapped[int] = mapped_column(nullable=False)
    biz_type: Mapped[str] = mapped_column(String(50), nullable=False)
    biz_id: Mapped[int | None] = mapped_column(nullable=True)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )
