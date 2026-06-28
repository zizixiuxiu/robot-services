from pydantic import BaseModel, ConfigDict
from datetime import datetime


class StockBalanceOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    product_id: int
    qty: int
    updated_at: datetime


class StockMovementOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    product_id: int
    change_qty: int
    before_qty: int
    after_qty: int
    biz_type: str
    biz_id: int | None
    remark: str | None
    created_at: datetime


class StockInCreate(BaseModel):
    product_id: int
    qty: int
    remark: str | None = None


class StockOutCreate(BaseModel):
    product_id: int
    qty: int
    remark: str | None = None
