from fastapi import APIRouter, Request, Depends, Form, status
from fastapi.responses import RedirectResponse
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from app.db import get_db
from app.services.orders import OrderService
from app.services.owners import OwnerService
from app.services.products import ProductService
from app.schemas.order import OrderCreate, OrderUpdate, OrderLineCreate
from datetime import date
from decimal import Decimal, InvalidOperation
from itertools import zip_longest

router = APIRouter(prefix="/orders")
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


def _build_validated_lines(
    line_product_id: list,
    line_qty: list,
    line_unit_price: list,
    line_substrate_category: list,
    line_surface_type: list,
    line_remark: list,
) -> tuple[list[OrderLineCreate], str | None]:
    """Build order lines from form data, validating that all fields are filled."""
    lines: list[OrderLineCreate] = []
    for pid, qty, price, sub, sur, rem in zip_longest(
        line_product_id,
        line_qty,
        line_unit_price,
        line_substrate_category,
        line_surface_type,
        line_remark,
        fillvalue="",
    ):
        if not pid or not qty:
            continue
        if not str(pid).strip() or not str(qty).strip():
            continue
        if int(qty) <= 0:
            continue
        if not str(price).strip():
            return [], "请填写所有明细行的单价。"
        if not str(sub).strip():
            return [], "请填写所有明细行的基材类别。"
        if not str(sur).strip():
            return [], "请选择所有明细行的单面/双面。"
        if not str(rem).strip():
            return [], "请填写所有明细行的备注。"
        unit_price = parse_money(price)
        lines.append(
            OrderLineCreate(
                product_id=int(pid),
                qty=int(qty),
                unit_price=unit_price,
                total_price=(Decimal(int(qty)) * unit_price).quantize(Decimal("0.01")),
                substrate_category=sub.strip(),
                surface_type=sur.strip(),
                remark=rem.strip() or None,
            )
        )
    return lines, None


def build_label_lines(order):
    label_lines = []
    for line in order.lines:
        product = line.product
        label_lines.append(
            {
                "spec": product.spec or "",
                "production_date": order.order_date,
                "substrate_category": line.substrate_category or "",
                "qty": line.qty,
                "product_color": product.product_name or "",
                "surface_type": line.surface_type or "",
            }
        )
    return label_lines


@router.get("/")
def list_orders(request: Request, error: str | None = None, db: Session = Depends(get_db)):
    orders = OrderService.list_orders(db)
    return templates.TemplateResponse(
        request, "orders/list.html", {"orders": orders, "error": error}
    )


@router.get("/new")
def new_order(request: Request, db: Session = Depends(get_db)):
    owners = OwnerService.list_owners(db)
    products = ProductService.list_products(db)
    return templates.TemplateResponse(
        request,
        "orders/form.html",
        {"order": None, "owners": owners, "products": products, "error": None},
    )


@router.post("/")
def create_order(
    request: Request,
    order_date: str = Form(),
    due_date: str | None = Form(default=None),
    owner_user_id: int = Form(),
    remark: str = Form(default=""),
    line_product_id: list[int] = Form(default_factory=list),
    line_qty: list[int] = Form(default_factory=list),
    line_unit_price: list[str] = Form(default_factory=list),
    line_substrate_category: list[str] = Form(default_factory=list),
    line_surface_type: list[str] = Form(default_factory=list),
    line_remark: list[str] = Form(default_factory=list),
    db: Session = Depends(get_db),
):
    owners = OwnerService.list_owners(db)
    products = ProductService.list_products(db)

    order_date_parsed = date.fromisoformat(order_date)
    due_date_parsed = date.fromisoformat(due_date) if due_date else None

    lines, error = _build_validated_lines(
        line_product_id,
        line_qty,
        line_unit_price,
        line_substrate_category,
        line_surface_type,
        line_remark,
    )
    if error:
        return templates.TemplateResponse(
            request,
            "orders/form.html",
            {
                "order": None,
                "owners": owners,
                "products": products,
                "error": error,
                "form": {
                    "order_date": order_date,
                    "due_date": due_date,
                    "owner_user_id": owner_user_id,
                    "remark": remark,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    if not lines:
        return templates.TemplateResponse(
            request,
            "orders/form.html",
            {
                "order": None,
                "owners": owners,
                "products": products,
                "error": "At least one line item is required.",
                "form": {
                    "order_date": order_date,
                    "due_date": due_date,
                    "owner_user_id": owner_user_id,
                    "remark": remark,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    order_no = OrderService.generate_order_no(db)
    data = OrderCreate(
        order_no=order_no,
        order_date=order_date_parsed,
        due_date=due_date_parsed,
        owner_user_id=owner_user_id,
        remark=remark or None,
        lines=lines,
    )
    OrderService.create_order(db, data)
    return RedirectResponse(url="/orders/", status_code=status.HTTP_303_SEE_OTHER)


@router.get("/{order_id}")
def order_detail(request: Request, order_id: int, db: Session = Depends(get_db)):
    order = OrderService.get_order(db, order_id)
    if not order:
        return RedirectResponse(url="/orders/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "orders/detail.html", {"order": order}
    )


@router.get("/{order_id}/print")
def order_print(request: Request, order_id: int, db: Session = Depends(get_db)):
    order = OrderService.get_order(db, order_id)
    if not order:
        return RedirectResponse(url="/orders/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "orders/print.html", {"order": order}
    )


@router.get("/{order_id}/labels")
def order_labels(request: Request, order_id: int, db: Session = Depends(get_db)):
    order = OrderService.get_order(db, order_id)
    if not order:
        return RedirectResponse(url="/orders/", status_code=status.HTTP_302_FOUND)
    response = templates.TemplateResponse(
        request,
        "orders/labels.html",
        {"order": order, "label_lines": build_label_lines(order)},
    )
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@router.post("/{order_id}/delete")
def delete_order(
    request: Request,
    order_id: int,
    db: Session = Depends(get_db),
):
    order = OrderService.get_order(db, order_id)
    if not order:
        return RedirectResponse(url="/orders/", status_code=status.HTTP_302_FOUND)
    try:
        OrderService.delete_order(db, order)
    except ValueError as exc:
        orders = OrderService.list_orders(db)
        return templates.TemplateResponse(
            request,
            "orders/list.html",
            {"orders": orders, "error": str(exc)},
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/orders/", status_code=status.HTTP_303_SEE_OTHER)


@router.get("/{order_id}/edit")
def edit_order(request: Request, order_id: int, db: Session = Depends(get_db)):
    order = OrderService.get_order(db, order_id)
    if not order:
        return RedirectResponse(url="/orders/", status_code=status.HTTP_302_FOUND)
    owners = OwnerService.list_owners(db)
    products = ProductService.list_products(db)
    return templates.TemplateResponse(
        request,
        "orders/form.html",
        {"order": order, "owners": owners, "products": products, "error": None},
    )


@router.post("/{order_id}/edit")
def update_order(
    request: Request,
    order_id: int,
    order_date: str = Form(),
    due_date: str | None = Form(default=None),
    owner_user_id: int = Form(),
    remark: str = Form(default=""),
    line_product_id: list[int] = Form(default_factory=list),
    line_qty: list[int] = Form(default_factory=list),
    line_unit_price: list[str] = Form(default_factory=list),
    line_substrate_category: list[str] = Form(default_factory=list),
    line_surface_type: list[str] = Form(default_factory=list),
    line_remark: list[str] = Form(default_factory=list),
    db: Session = Depends(get_db),
):
    order = OrderService.get_order(db, order_id)
    if not order:
        return RedirectResponse(url="/orders/", status_code=status.HTTP_302_FOUND)

    owners = OwnerService.list_owners(db)
    products = ProductService.list_products(db)

    order_date_parsed = date.fromisoformat(order_date)
    due_date_parsed = date.fromisoformat(due_date) if due_date else None

    lines, error = _build_validated_lines(
        line_product_id,
        line_qty,
        line_unit_price,
        line_substrate_category,
        line_surface_type,
        line_remark,
    )
    if error:
        return templates.TemplateResponse(
            request,
            "orders/form.html",
            {
                "order": order,
                "owners": owners,
                "products": products,
                "error": error,
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    if not lines:
        return templates.TemplateResponse(
            request,
            "orders/form.html",
            {
                "order": order,
                "owners": owners,
                "products": products,
                "error": "At least one line item is required.",
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    data = OrderUpdate(
        order_date=order_date_parsed,
        due_date=due_date_parsed,
        owner_user_id=owner_user_id,
        remark=remark or None,
        lines=lines,
    )
    OrderService.update_order(db, order, data)
    return RedirectResponse(url=f"/orders/{order_id}", status_code=status.HTTP_303_SEE_OTHER)
