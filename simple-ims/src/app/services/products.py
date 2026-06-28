from sqlalchemy.orm import Session
from sqlalchemy.exc import IntegrityError
from sqlalchemy import cast, Integer
from app.models.inbound_order_line import InboundOrderLine
from app.models.order_line import OrderLine
from app.models.outbound_order_line import OutboundOrderLine
from app.models.product import Product
from app.models.stock_balance import StockBalance
from app.models.stock_movement import StockMovement
from app.schemas.product import ProductCreate, ProductUpdate


class ProductService:
    PRODUCT_CODE_START = 1000001

    @staticmethod
    def list_products(db: Session):
        return db.query(Product).order_by(Product.id).all()

    @staticmethod
    def get_product(db: Session, product_id: int):
        return db.query(Product).filter(Product.id == product_id).first()

    @staticmethod
    def get_product_by_code(db: Session, product_code: str):
        return db.query(Product).filter(Product.product_code == product_code).first()

    @staticmethod
    def get_product_by_name_and_spec(
        db: Session,
        product_name: str,
        spec: str | None,
        exclude_product_id: int | None = None,
    ):
        normalized_name = (product_name or "").strip()
        normalized_spec = (spec or "").strip()
        query = db.query(Product).filter(
            Product.product_name == normalized_name,
            Product.spec == normalized_spec,
        )
        if exclude_product_id is not None:
            query = query.filter(Product.id != exclude_product_id)
        return query.first()

    @staticmethod
    def generate_product_code(db: Session) -> str:
        max_numeric_code = (
            db.query(cast(Product.product_code, Integer))
            .filter(Product.product_code.op("GLOB")("[0-9]*"))
            .order_by(cast(Product.product_code, Integer).desc())
            .limit(1)
            .scalar()
        )
        next_code = ProductService.PRODUCT_CODE_START if max_numeric_code is None else max(max_numeric_code + 1, ProductService.PRODUCT_CODE_START)
        return str(next_code)

    @staticmethod
    def create_product(db: Session, data: ProductCreate):
        payload = data.model_dump()
        payload["product_name"] = (payload.get("product_name") or "").strip()
        payload["spec"] = (payload.get("spec") or "").strip()
        payload["unit"] = (payload.get("unit") or "").strip()
        payload["substrate_category"] = (payload.get("substrate_category") or "").strip()
        payload["surface_type"] = (payload.get("surface_type") or "").strip()
        if ProductService.get_product_by_name_and_spec(
            db, payload["product_name"], payload["spec"]
        ):
            raise ValueError("相同的产品名称和规格已存在，不能重复创建。")
        if not payload.get("product_code"):
            payload["product_code"] = ProductService.generate_product_code(db)
        product = Product(**payload)
        db.add(product)
        try:
            db.commit()
            db.refresh(product)
        except IntegrityError:
            db.rollback()
            raise
        return product

    @staticmethod
    def delete_product(db: Session, product: Product):
        usages = [
            db.query(OrderLine).filter(OrderLine.product_id == product.id).count(),
            db.query(InboundOrderLine).filter(InboundOrderLine.product_id == product.id).count(),
            db.query(OutboundOrderLine).filter(OutboundOrderLine.product_id == product.id).count(),
            db.query(StockMovement).filter(StockMovement.product_id == product.id).count(),
        ]
        balance = db.query(StockBalance).filter(StockBalance.product_id == product.id).first()
        if any(usages) or (balance and balance.qty != 0):
            raise ValueError("该产品已有订单、库存或流水记录，不能直接删除。")
        if balance:
            db.delete(balance)
        db.delete(product)
        db.commit()

    @staticmethod
    def update_product(db: Session, product: Product, data: ProductUpdate):
        updates = data.model_dump(exclude_unset=True)
        if "product_name" in updates:
            updates["product_name"] = (updates.get("product_name") or "").strip()
        if "spec" in updates:
            updates["spec"] = (updates.get("spec") or "").strip()
        if "unit" in updates:
            updates["unit"] = (updates.get("unit") or "").strip()
        if "substrate_category" in updates:
            updates["substrate_category"] = (updates.get("substrate_category") or "").strip()
        if "surface_type" in updates:
            updates["surface_type"] = (updates.get("surface_type") or "").strip()

        candidate_name = updates.get("product_name", product.product_name)
        candidate_spec = updates.get("spec", product.spec)
        if ProductService.get_product_by_name_and_spec(
            db, candidate_name, candidate_spec, exclude_product_id=product.id
        ):
            raise ValueError("相同的产品名称和规格已存在，不能重复创建。")

        for field, value in updates.items():
            setattr(product, field, value)
        try:
            db.commit()
            db.refresh(product)
        except IntegrityError:
            db.rollback()
            raise
        return product
