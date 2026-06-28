from fastapi import APIRouter, Request, Depends, Form, status
from fastapi.responses import RedirectResponse
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from app.db import get_db
from app.services.inbounds import InboundService
from app.services.orders import OrderService
from app.services.products import ProductService
from app.schemas.inbound import InboundOrderCreate, InboundOrderLineCreate
from datetime import date
from itertools import zip_longest

router = APIRouter(prefix="/inbounds")
templates = create_templates()


@router.get("/")
def list_inbounds(request: Request, error: str | None = None, db: Session = Depends(get_db)):
    inbounds = InboundService.list_inbounds(db)
    return templates.TemplateResponse(
        request, "inbounds/list.html", {"inbounds": inbounds, "error": error}
    )


@router.get("/new")
def new_inbound(request: Request, db: Session = Depends(get_db)):
    products = ProductService.list_products(db)
    orders = OrderService.list_orders_for_inbound(db)
    return templates.TemplateResponse(
        request,
        "inbounds/form.html",
        {"inbound": None, "products": products, "orders": orders, "error": None},
    )


@router.post("/")
def create_inbound(
    request: Request,
    inbound_date: str = Form(),
    related_order_id: str = Form(default=""),
    remark: str = Form(default=""),
    line_product_id: list[int] = Form(default_factory=list),
    line_qty: list[int] = Form(default_factory=list),
    line_remark: list[str] = Form(default_factory=list),
    db: Session = Depends(get_db),
):
    products = ProductService.list_products(db)
    orders = OrderService.list_orders_for_inbound(db)

    inbound_date_parsed = date.fromisoformat(inbound_date)
    related_order_id_parsed = int(related_order_id) if related_order_id else None

    lines = []
    for pid, qty, rem in zip_longest(line_product_id, line_qty, line_remark, fillvalue=""):
        if pid and qty and int(qty) > 0:
            lines.append(InboundOrderLineCreate(product_id=int(pid), qty=int(qty), remark=rem or None))

    if not lines:
        return templates.TemplateResponse(
            request,
            "inbounds/form.html",
            {
                "inbound": None,
                "products": products,
                "orders": orders,
                "error": "At least one line item is required.",
                "form": {
                    "inbound_date": inbound_date,
                    "related_order_id": related_order_id_parsed,
                    "remark": remark,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    inbound_no = InboundService.generate_inbound_no(db)
    data = InboundOrderCreate(
        inbound_no=inbound_no,
        inbound_date=inbound_date_parsed,
        related_order_id=related_order_id_parsed,
        remark=remark or None,
        lines=lines,
    )
    try:
        InboundService.create_inbound(db, data)
    except ValueError as e:
        return templates.TemplateResponse(
            request,
            "inbounds/form.html",
            {
                "inbound": None,
                "products": products,
                "orders": orders,
                "error": str(e),
                "form": {
                    "inbound_date": inbound_date,
                    "related_order_id": related_order_id_parsed,
                    "remark": remark,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/inbounds/", status_code=status.HTTP_303_SEE_OTHER)


@router.get("/{inbound_id}")
def inbound_detail(request: Request, inbound_id: int, db: Session = Depends(get_db)):
    inbound = InboundService.get_inbound(db, inbound_id)
    if not inbound:
        return RedirectResponse(url="/inbounds/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "inbounds/detail.html", {"inbound": inbound}
    )


@router.get("/{inbound_id}/print")
def inbound_print(request: Request, inbound_id: int, db: Session = Depends(get_db)):
    inbound = InboundService.get_inbound(db, inbound_id)
    if not inbound:
        return RedirectResponse(url="/inbounds/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "inbounds/print.html", {"inbound": inbound}
    )


@router.post("/{inbound_id}/delete")
def delete_inbound(
    request: Request,
    inbound_id: int,
    db: Session = Depends(get_db),
):
    inbound = InboundService.get_inbound(db, inbound_id)
    if not inbound:
        return RedirectResponse(url="/inbounds/", status_code=status.HTTP_302_FOUND)
    try:
        InboundService.delete_inbound(db, inbound)
    except ValueError as exc:
        inbounds = InboundService.list_inbounds(db)
        return templates.TemplateResponse(
            request,
            "inbounds/list.html",
            {"inbounds": inbounds, "error": str(exc)},
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/inbounds/", status_code=status.HTTP_303_SEE_OTHER)
