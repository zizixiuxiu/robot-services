from pydantic import BaseModel, ConfigDict
from datetime import datetime


class OwnerBase(BaseModel):
    name: str
    code: str
    enabled: bool = True


class OwnerCreate(OwnerBase):
    pass


class OwnerUpdate(BaseModel):
    name: str | None = None
    code: str | None = None
    enabled: bool | None = None


class OwnerOut(OwnerBase):
    model_config = ConfigDict(from_attributes=True)

    id: int
    created_at: datetime
    updated_at: datetime
