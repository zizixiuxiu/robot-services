from pydantic import BaseModel, ConfigDict
from datetime import datetime, date


class InboundOrderLineBase(BaseModel):
    product_id: int
    qty: int
    remark: str | None = None


class InboundOrderLineCreate(InboundOrderLineBase):
    pass


class InboundOrderLineOut(InboundOrderLineBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    inbound_order_id: int


class InboundOrderBase(BaseModel):
    inbound_no: str
    inbound_date: date
    related_order_id: int | None = None
    remark: str | None = None


class InboundOrderCreate(BaseModel):
    inbound_no: str
    inbound_date: date
    related_order_id: int | None = None
    remark: str | None = None
    lines: list[InboundOrderLineCreate]


class InboundOrderOut(InboundOrderBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    created_at: datetime
    lines: list[InboundOrderLineOut]
