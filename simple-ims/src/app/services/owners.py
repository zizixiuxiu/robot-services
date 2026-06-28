from sqlalchemy.orm import Session
from app.models.order import Order
from app.models.outbound_order import OutboundOrder
from app.models.owner import Owner
from app.schemas.owner import OwnerCreate, OwnerUpdate


class OwnerService:
    @staticmethod
    def seed_defaults(db: Session) -> None:
        defaults = [
            {"name": "奢匠", "code": "001"},
            {"name": "固豪", "code": "002"},
            {"name": "PVC车间", "code": "003"},
            {"name": "久派", "code": "004"},
        ]
        for data in defaults:
            existing = db.query(Owner).filter(Owner.code == data["code"]).first()
            if not existing:
                owner = Owner(name=data["name"], code=data["code"])
                db.add(owner)
            elif existing.name != data["name"]:
                existing.name = data["name"]
        db.commit()

    @staticmethod
    def list_owners(db: Session):
        return db.query(Owner).order_by(Owner.id).all()

    @staticmethod
    def get_owner(db: Session, owner_id: int):
        return db.query(Owner).filter(Owner.id == owner_id).first()

    @staticmethod
    def create_owner(db: Session, data: OwnerCreate):
        owner = Owner(**data.model_dump())
        db.add(owner)
        db.commit()
        db.refresh(owner)
        return owner

    @staticmethod
    def update_owner(db: Session, owner: Owner, data: OwnerUpdate):
        updates = data.model_dump(exclude_unset=True)
        for field, value in updates.items():
            setattr(owner, field, value)
        db.commit()
        db.refresh(owner)
        return owner

    @staticmethod
    def delete_owner(db: Session, owner: Owner):
        order_count = db.query(Order).filter(Order.owner_user_id == owner.id).count()
        outbound_count = db.query(OutboundOrder).filter(OutboundOrder.owner_user_id == owner.id).count()
        if order_count or outbound_count:
            raise ValueError("该客户已有订单或出库记录，不能直接删除。")
        db.delete(owner)
        db.commit()
