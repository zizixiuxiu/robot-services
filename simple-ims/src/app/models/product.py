from datetime import datetime
from decimal import Decimal
from sqlalchemy import String, Boolean, DateTime, func, Numeric, Text
from sqlalchemy.orm import Mapped, mapped_column
from app.db import Base


class Product(Base):
    __tablename__ = "products"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    product_code: Mapped[str] = mapped_column(
        String(255), nullable=False, unique=True, index=True
    )
    product_name: Mapped[str] = mapped_column(String(255), nullable=False)
    spec: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    unit: Mapped[str] = mapped_column(String(50), nullable=False, default="")
    unit_price: Mapped[Decimal] = mapped_column(Numeric(12, 2), nullable=False, default=0)
    substrate_category: Mapped[str] = mapped_column(Text, nullable=False, default="")
    surface_type: Mapped[str] = mapped_column(Text, nullable=False, default="")
    enabled: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)
    remark: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), nullable=False
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        server_default=func.now(),
        onupdate=func.now(),
        nullable=False,
    )
