from fastapi import APIRouter, Request, Depends, Form, status
from fastapi.responses import RedirectResponse
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from app.db import get_db
from app.services.outbounds import OutboundService
from app.services.owners import OwnerService
from app.services.products import ProductService
from app.services.orders import OrderService
from app.schemas.outbound import OutboundOrderCreate, OutboundOrderLineCreate
from datetime import date
from decimal import Decimal, InvalidOperation
from itertools import zip_longest

router = APIRouter(prefix="/outbounds")
templates = create_templates()


def parse_money(value: str | None) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        amount = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return Decimal("0")
    if amount < 0:
        return Decimal("0")
    return amount.quantize(Decimal("0.01"))


@router.get("/")
def list_outbounds(request: Request, error: str | None = None, db: Session = Depends(get_db)):
    outbounds = OutboundService.list_outbounds(db)
    return templates.TemplateResponse(
        request, "outbounds/list.html", {"outbounds": outbounds, "error": error}
    )


@router.get("/new")
def new_outbound(request: Request, db: Session = Depends(get_db)):
    owners = OwnerService.list_owners(db)
    products = ProductService.list_products(db)
    orders = OrderService.list_orders_for_outbound(db)
    return templates.TemplateResponse(
        request,
        "outbounds/form.html",
        {"outbound": None, "owners": owners, "products": products, "orders": orders, "error": None},
    )


@router.post("/")
def create_outbound(
    request: Request,
    outbound_date: str = Form(),
    owner_user_id: int = Form(),
    related_inbound_id: str = Form(default=""),
    related_order_id: str = Form(default=""),
    remark: str = Form(default=""),
    line_product_id: list[int] = Form(default_factory=list),
    line_qty: list[int] = Form(default_factory=list),
    line_unit_price: list[str] = Form(default_factory=list),
    line_remark: list[str] = Form(default_factory=list),
    db: Session = Depends(get_db),
):
    owners = OwnerService.list_owners(db)
    products = ProductService.list_products(db)
    orders = OrderService.list_orders_for_outbound(db)

    outbound_date_parsed = date.fromisoformat(outbound_date)
    related_inbound_id_parsed = int(related_inbound_id) if related_inbound_id else None
    related_order_id_parsed = int(related_order_id) if related_order_id else None

    lines = []
    for pid, qty, price, rem in zip_longest(line_product_id, line_qty, line_unit_price, line_remark, fillvalue=""):
        if pid and qty and int(qty) > 0:
            unit_price = parse_money(price)
            lines.append(
                OutboundOrderLineCreate(
                    product_id=int(pid),
                    qty=int(qty),
                    unit_price=unit_price,
                    total_price=(Decimal(int(qty)) * unit_price).quantize(Decimal("0.01")),
                    remark=rem or None,
                )
            )

    if not lines:
        return templates.TemplateResponse(
            request,
            "outbounds/form.html",
            {
                "outbound": None,
                "owners": owners,
                "products": products,
                "orders": orders,
                "error": "At least one line item is required.",
                "form": {
                    "outbound_date": outbound_date,
                    "owner_user_id": owner_user_id,
                    "related_order_id": related_order_id_parsed,
                    "related_inbound_id": related_inbound_id_parsed,
                    "remark": remark,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    outbound_no = OutboundService.generate_outbound_no(db)
    data = OutboundOrderCreate(
        outbound_no=outbound_no,
        outbound_date=outbound_date_parsed,
        owner_user_id=owner_user_id,
        related_order_id=related_order_id_parsed,
        related_inbound_id=related_inbound_id_parsed,
        remark=remark or None,
        lines=lines,
    )
    try:
        OutboundService.create_outbound(db, data)
    except ValueError as e:
        return templates.TemplateResponse(
            request,
            "outbounds/form.html",
            {
                "outbound": None,
                "owners": owners,
                "products": products,
                "orders": orders,
                "error": str(e),
                "form": {
                    "outbound_date": outbound_date,
                    "owner_user_id": owner_user_id,
                    "related_order_id": related_order_id_parsed,
                    "related_inbound_id": related_inbound_id_parsed,
                    "remark": remark,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/outbounds/", status_code=status.HTTP_303_SEE_OTHER)


@router.get("/{outbound_id}")
def outbound_detail(request: Request, outbound_id: int, db: Session = Depends(get_db)):
    outbound = OutboundService.get_outbound(db, outbound_id)
    if not outbound:
        return RedirectResponse(url="/outbounds/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "outbounds/detail.html", {"outbound": outbound}
    )


@router.get("/{outbound_id}/print")
def outbound_print(request: Request, outbound_id: int, db: Session = Depends(get_db)):
    outbound = OutboundService.get_outbound(db, outbound_id)
    if not outbound:
        return RedirectResponse(url="/outbounds/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "outbounds/print.html", {"outbound": outbound}
    )


@router.post("/{outbound_id}/delete")
def delete_outbound(
    request: Request,
    outbound_id: int,
    db: Session = Depends(get_db),
):
    outbound = OutboundService.get_outbound(db, outbound_id)
    if not outbound:
        return RedirectResponse(url="/outbounds/", status_code=status.HTTP_302_FOUND)
    try:
        OutboundService.delete_outbound(db, outbound)
    except ValueError as exc:
        outbounds = OutboundService.list_outbounds(db)
        return templates.TemplateResponse(
            request,
            "outbounds/list.html",
            {"outbounds": outbounds, "error": str(exc)},
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/outbounds/", status_code=status.HTTP_303_SEE_OTHER)
