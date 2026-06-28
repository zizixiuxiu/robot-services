from fastapi import APIRouter, Request
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from app.db import SessionLocal
from app.models.owner import Owner
from app.models.product import Product
from app.models.stock_balance import StockBalance
from app.models.order import Order
from app.models.inbound_order import InboundOrder
from app.models.outbound_order import OutboundOrder

router = APIRouter()

templates = create_templates()


@router.get("/")
def home(request: Request):
    db: Session = SessionLocal()
    try:
        counts = {
            "owners": db.query(Owner).count(),
            "products": db.query(Product).count(),
            "stock": db.query(StockBalance).count(),
            "orders": db.query(Order).count(),
            "inbounds": db.query(InboundOrder).count(),
            "outbounds": db.query(OutboundOrder).count(),
        }
    finally:
        db.close()
    return templates.TemplateResponse(request, "home.html", {"counts": counts})
