from __future__ import annotations

import os
import tempfile
from datetime import date
from decimal import Decimal
from typing import Any

from fastapi import APIRouter, Depends, File, Form, Request, UploadFile
from sqlalchemy.orm import Session

from app.ai.matching import ensure_product_for_item
from app.ai.recognizers import BaseRecognizer, get_recognizer
from app.db import get_db
from app.schemas.order import OrderCreate, OrderLineCreate
from app.services.orders import OrderService
from app.services.owners import OwnerService
from app.web.templating import create_templates

router = APIRouter(prefix="/ai")
templates = create_templates()


def _save_upload(image: UploadFile) -> str:
    suffix = os.path.splitext(image.filename or ".png")[1] or ".png"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(image.file.read())
        return tmp.name


@router.post("/orders/create")
def create_order_from_image(
    image: UploadFile = File(...),
    owner_user_id: int = Form(default=0),
    due_date: str | None = Form(default=None),
    recognizer: BaseRecognizer = Depends(get_recognizer),
    db: Session = Depends(get_db),
):
    """Upload an image, recognize it, match/create products, and create an order."""
    tmp_path = _save_upload(image)
    try:
        result = recognizer.recognize(tmp_path)
    finally:
        os.unlink(tmp_path)

    if not result.items:
        return {"ok": False, "message": "未识别到订单明细，无法创建订单。"}

    owner = None
    if owner_user_id:
        owner = OwnerService.get_owner(db, owner_user_id)
    if not owner:
        first_owner = OwnerService.list_owners(db)
        if first_owner:
            owner = first_owner[0]
            owner_user_id = owner.id
        else:
            return {"ok": False, "message": "没有可用客户信息，请先创建客户。"}

    order_date = date.today()
    if result.date:
        try:
            order_date = date.fromisoformat(result.date)
        except ValueError:
            pass

    due_date_parsed = None
    if due_date:
        try:
            due_date_parsed = date.fromisoformat(due_date)
        except ValueError:
            pass

    lines: list[OrderLineCreate] = []
    created_products: list[dict[str, Any]] = []
    matched_products: list[dict[str, Any]] = []
    warnings: list[str] = []

    for item in result.items:
        match = ensure_product_for_item(
            db,
            name=item.name,
            spec=item.spec,
            unit=item.unit or "张",
            unit_price=Decimal(str(item.qty)) if item.qty else Decimal("0"),
            substrate_category="",
            surface_type="",
            remark=item.remark,
            min_score=0.75,
        )

        if not match.product:
            warnings.append(match.warning or f"无法处理产品：{item.name}")
            continue

        if match.created_new:
            created_products.append(
                {
                    "product_id": match.product.id,
                    "product_code": match.product.product_code,
                    "product_name": match.product.product_name,
                    "spec": match.product.spec,
                }
            )
        else:
            matched_products.append(
                {
                    "product_id": match.product.id,
                    "product_code": match.product.product_code,
                    "product_name": match.product.product_name,
                    "spec": match.product.spec,
                    "match_score": match.score,
                }
            )

        unit_price = match.product.unit_price or Decimal("0")
        qty = int(item.qty) if item.qty else 1
        lines.append(
            OrderLineCreate(
                product_id=match.product.id,
                qty=qty,
                unit_price=unit_price,
                total_price=(Decimal(qty) * unit_price).quantize(Decimal("0.01")),
                substrate_category=match.product.substrate_category or "",
                surface_type=match.product.surface_type or "",
                remark=item.remark or None,
            )
        )

    if not lines:
        return {"ok": False, "message": "没有可创建的订单明细。", "warnings": warnings}

    order_data = OrderCreate(
        order_no=OrderService.generate_order_no(db),
        order_date=order_date,
        due_date=due_date_parsed,
        owner_user_id=owner_user_id,
        remark=result.remark or result.description or None,
        lines=lines,
    )
    order = OrderService.create_order(db, order_data)

    return {
        "ok": True,
        "order_id": order.id,
        "order_no": order.order_no,
        "order_date": str(order.order_date),
        "owner": owner.name if owner else None,
        "remark": result.remark or result.description or None,
        "items_count": len(lines),
        "created_products": created_products,
        "matched_products": matched_products,
        "warnings": warnings,
    }


@router.get("/demo")
def demo_page(request: Request, db: Session = Depends(get_db)):
    """Browser-friendly upload page for testing AI recognition."""
    owners = OwnerService.list_owners(db)
    return templates.TemplateResponse(
        request,
        "ai/demo.html",
        {
            "path": request.url.path,
            "owners": owners,
        },
    )
