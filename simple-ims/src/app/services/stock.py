from sqlalchemy.orm import Session
from app.models.stock_balance import StockBalance
from app.models.stock_movement import StockMovement
from app.schemas.stock import StockInCreate, StockOutCreate


class StockService:
    @staticmethod
    def list_balances(db: Session):
        return (
            db.query(StockBalance)
            .order_by(StockBalance.product_id)
            .all()
        )

    @staticmethod
    def get_balance_by_product(db: Session, product_id: int):
        return (
            db.query(StockBalance)
            .filter(StockBalance.product_id == product_id)
            .first()
        )

    @staticmethod
    def list_movements(db: Session, product_id: int | None = None):
        query = db.query(StockMovement).order_by(StockMovement.id.desc())
        if product_id is not None:
            query = query.filter(StockMovement.product_id == product_id)
        return query.all()

    @staticmethod
    def stock_in(db: Session, data: StockInCreate):
        balance = StockService.get_balance_by_product(db, data.product_id)
        before_qty = balance.qty if balance else 0
        after_qty = before_qty + data.qty

        if balance:
            balance.qty = after_qty
        else:
            balance = StockBalance(
                product_id=data.product_id,
                qty=after_qty,
            )
            db.add(balance)

        movement = StockMovement(
            product_id=data.product_id,
            change_qty=data.qty,
            before_qty=before_qty,
            after_qty=after_qty,
            biz_type="manual_in",
            biz_id=None,
            remark=data.remark,
        )
        db.add(movement)
        db.commit()
        db.refresh(balance)
        db.refresh(movement)
        return balance, movement

    @staticmethod
    def stock_in_with_biz(db: Session, product_id: int, qty: int, remark: str | None, biz_type: str, biz_id: int | None):
        balance = StockService.get_balance_by_product(db, product_id)
        before_qty = balance.qty if balance else 0
        after_qty = before_qty + qty

        if balance:
            balance.qty = after_qty
        else:
            balance = StockBalance(
                product_id=product_id,
                qty=after_qty,
            )
            db.add(balance)

        movement = StockMovement(
            product_id=product_id,
            change_qty=qty,
            before_qty=before_qty,
            after_qty=after_qty,
            biz_type=biz_type,
            biz_id=biz_id,
            remark=remark,
        )
        db.add(movement)
        db.commit()
        db.refresh(balance)
        db.refresh(movement)
        return balance, movement

    @staticmethod
    def stock_out(db: Session, data: StockOutCreate):
        balance = StockService.get_balance_by_product(db, data.product_id)
        before_qty = balance.qty if balance else 0

        if before_qty < data.qty:
            raise ValueError("Insufficient stock")

        after_qty = before_qty - data.qty

        if balance:
            balance.qty = after_qty
        else:
            balance = StockBalance(
                product_id=data.product_id,
                qty=after_qty,
            )
            db.add(balance)

        movement = StockMovement(
            product_id=data.product_id,
            change_qty=-data.qty,
            before_qty=before_qty,
            after_qty=after_qty,
            biz_type="manual_out",
            biz_id=None,
            remark=data.remark,
        )
        db.add(movement)
        db.commit()
        db.refresh(balance)
        db.refresh(movement)
        return balance, movement

    @staticmethod
    def stock_out_with_biz(db: Session, product_id: int, qty: int, remark: str | None, biz_type: str, biz_id: int | None):
        balance = StockService.get_balance_by_product(db, product_id)
        before_qty = balance.qty if balance else 0

        if before_qty < qty:
            raise ValueError("Insufficient stock")

        after_qty = before_qty - qty

        if balance:
            balance.qty = after_qty
        else:
            balance = StockBalance(
                product_id=product_id,
                qty=after_qty,
            )
            db.add(balance)

        movement = StockMovement(
            product_id=product_id,
            change_qty=-qty,
            before_qty=before_qty,
            after_qty=after_qty,
            biz_type=biz_type,
            biz_id=biz_id,
            remark=remark,
        )
        db.add(movement)
        db.commit()
        db.refresh(balance)
        db.refresh(movement)
        return balance, movement
