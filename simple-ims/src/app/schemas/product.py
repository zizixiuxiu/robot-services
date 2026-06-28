from pydantic import BaseModel, ConfigDict
from datetime import datetime
from decimal import Decimal


class ProductBase(BaseModel):
    product_code: str
    product_name: str
    spec: str = ""
    unit: str = ""
    unit_price: Decimal = Decimal("0")
    substrate_category: str = ""
    surface_type: str = ""
    enabled: bool = True
    remark: str | None = None


class ProductCreate(ProductBase):
    pass


class ProductQuickCreate(BaseModel):
    product_name: str
    spec: str = ""
    unit: str = ""
    unit_price: Decimal = Decimal("0")
    substrate_category: str = ""
    surface_type: str = ""
    remark: str | None = None


class ProductUpdate(BaseModel):
    product_code: str | None = None
    product_name: str | None = None
    spec: str | None = None
    unit: str | None = None
    unit_price: Decimal | None = None
    substrate_category: str | None = None
    surface_type: str | None = None
    enabled: bool | None = None
    remark: str | None = None


class ProductOut(ProductBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    created_at: datetime
    updated_at: datetime
