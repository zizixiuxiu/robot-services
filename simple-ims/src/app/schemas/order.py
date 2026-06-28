from pydantic import BaseModel, ConfigDict
from datetime import datetime, date
from decimal import Decimal


class OrderLineBase(BaseModel):
    product_id: int
    qty: int
    unit_price: Decimal = Decimal("0")
    total_price: Decimal = Decimal("0")
    substrate_category: str = ""
    surface_type: str = ""
    remark: str | None = None


class OrderLineCreate(OrderLineBase):
    pass


class OrderLineOut(OrderLineBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    order_id: int
    inbound_qty: int
    shipped_qty: int


class OrderBase(BaseModel):
    order_no: str
    order_date: date
    due_date: date | None = None
    owner_user_id: int
    status: str = "producing"
    remark: str | None = None


class OrderCreate(BaseModel):
    order_no: str
    order_date: date
    due_date: date | None = None
    owner_user_id: int
    remark: str | None = None
    lines: list[OrderLineCreate]


class OrderUpdate(BaseModel):
    order_date: date | None = None
    due_date: date | None = None
    owner_user_id: int | None = None
    remark: str | None = None
    lines: list[OrderLineCreate] | None = None


class OrderOut(OrderBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    created_at: datetime
    updated_at: datetime
    lines: list[OrderLineOut]
