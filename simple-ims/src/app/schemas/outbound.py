from pydantic import BaseModel, ConfigDict
from datetime import datetime, date
from decimal import Decimal


class OutboundOrderLineBase(BaseModel):
    product_id: int
    qty: int
    unit_price: Decimal = Decimal("0")
    total_price: Decimal = Decimal("0")
    remark: str | None = None


class OutboundOrderLineCreate(OutboundOrderLineBase):
    pass


class OutboundOrderLineOut(OutboundOrderLineBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    outbound_order_id: int


class OutboundOrderBase(BaseModel):
    outbound_no: str
    outbound_date: date
    owner_user_id: int
    related_order_id: int | None = None
    related_inbound_id: int | None = None
    remark: str | None = None


class OutboundOrderCreate(BaseModel):
    outbound_no: str
    outbound_date: date
    owner_user_id: int
    related_order_id: int | None = None
    related_inbound_id: int | None = None
    remark: str | None = None
    lines: list[OutboundOrderLineCreate]


class OutboundOrderOut(OutboundOrderBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    created_at: datetime
    lines: list[OutboundOrderLineOut]
