from fastapi import APIRouter, Request, Depends, Form, status
from fastapi.responses import RedirectResponse
from fastapi.responses import JSONResponse
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from sqlalchemy.exc import IntegrityError
from app.db import get_db
from app.services.products import ProductService
from app.schemas.product import ProductCreate, ProductUpdate, ProductQuickCreate
from decimal import Decimal, InvalidOperation


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


router = APIRouter(prefix="/products")
templates = create_templates()


@router.get("/")
def list_products(request: Request, error: str | None = None, db: Session = Depends(get_db)):
    products = ProductService.list_products(db)
    return templates.TemplateResponse(
        request, "products/list.html", {"products": products, "error": error}
    )


@router.get("/new")
def new_product(request: Request):
    return templates.TemplateResponse(
        request,
        "products/form.html",
        {"product": None, "error": None, "next_product_code": None},
    )


@router.post("/")
def create_product(
    request: Request,
    product_name: str = Form(),
    spec: str = Form(default=""),
    unit: str = Form(default=""),
    unit_price: str = Form(default="0"),
    substrate_category: str = Form(default=""),
    surface_type: str = Form(default=""),
    enabled: bool = Form(default=True),
    remark: str = Form(default=""),
    db: Session = Depends(get_db),
):
    generated_code = ProductService.generate_product_code(db)
    data = ProductCreate(
        product_code=generated_code,
        product_name=product_name,
        spec=spec,
        unit=unit,
        unit_price=parse_money(unit_price),
        substrate_category=substrate_category,
        surface_type=surface_type,
        enabled=enabled,
        remark=remark or None,
    )
    try:
        ProductService.create_product(db, data)
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "products/form.html",
            {
                "product": data,
                "next_product_code": generated_code,
                "error": str(exc),
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    except IntegrityError:
        return templates.TemplateResponse(
            request,
            "products/form.html",
            {
                "product": data,
                "next_product_code": generated_code,
                "error": f"Product code '{generated_code}' already exists.",
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/products/", status_code=status.HTTP_303_SEE_OTHER)


@router.post("/{product_id}/delete")
def delete_product(
    request: Request,
    product_id: int,
    db: Session = Depends(get_db),
):
    product = ProductService.get_product(db, product_id)
    if not product:
        return RedirectResponse(url="/products/", status_code=status.HTTP_302_FOUND)
    try:
        ProductService.delete_product(db, product)
    except ValueError as exc:
        products = ProductService.list_products(db)
        return templates.TemplateResponse(
            request,
            "products/list.html",
            {"products": products, "error": str(exc)},
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/products/", status_code=status.HTTP_303_SEE_OTHER)


@router.post("/quick-create")
def quick_create_product(
    payload: ProductQuickCreate,
    db: Session = Depends(get_db),
):
    data = ProductCreate(
        product_code=ProductService.generate_product_code(db),
        product_name=payload.product_name,
        spec=payload.spec,
        unit=payload.unit,
        unit_price=parse_money(payload.unit_price),
        substrate_category=payload.substrate_category,
        surface_type=payload.surface_type,
        enabled=True,
        remark=payload.remark,
    )
    try:
        product = ProductService.create_product(db, data)
        created = True
    except ValueError:
        product = ProductService.get_product_by_name_and_spec(
            db, payload.product_name, payload.spec
        )
        if not product:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content={"ok": False, "message": "产品快速创建失败，请检查名称和规格。"},
            )
        created = False

    return {
        "ok": True,
        "created": created,
        "product": {
            "id": product.id,
            "product_code": product.product_code,
            "product_name": product.product_name,
            "spec": product.spec,
            "unit": product.unit,
            "unit_price": str(product.unit_price or 0),
            "substrate_category": product.substrate_category,
            "surface_type": product.surface_type,
            "label": f"{product.product_code} - {product.product_name}",
        },
        "unit_price": str(product.unit_price or 0)
    }


@router.get("/{product_id}")
def product_detail(request: Request, product_id: int, db: Session = Depends(get_db)):
    product = ProductService.get_product(db, product_id)
    if not product:
        return RedirectResponse(url="/products/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "products/detail.html", {"product": product}
    )


@router.get("/{product_id}/edit")
def edit_product(request: Request, product_id: int, db: Session = Depends(get_db)):
    product = ProductService.get_product(db, product_id)
    if not product:
        return RedirectResponse(url="/products/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "products/form.html", {"product": product, "error": None, "next_product_code": None}
    )


@router.post("/{product_id}/edit")
def update_product(
    request: Request,
    product_id: int,
    product_code: str = Form(),
    product_name: str = Form(),
    spec: str = Form(default=""),
    unit: str = Form(default=""),
    unit_price: str = Form(default="0"),
    substrate_category: str = Form(default=""),
    surface_type: str = Form(default=""),
    enabled: bool = Form(default=True),
    remark: str = Form(default=""),
    db: Session = Depends(get_db),
):
    product = ProductService.get_product(db, product_id)
    if not product:
        return RedirectResponse(url="/products/", status_code=status.HTTP_302_FOUND)
    data = ProductUpdate(
        product_code=product_code,
        product_name=product_name,
        spec=spec,
        unit=unit,
        unit_price=parse_money(unit_price),
        substrate_category=substrate_category,
        surface_type=surface_type,
        enabled=enabled,
        remark=remark or None,
    )
    try:
        ProductService.update_product(db, product, data)
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "products/form.html",
            {
                "product": data,
                "next_product_code": None,
                "error": str(exc),
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    except IntegrityError:
        return templates.TemplateResponse(
            request,
            "products/form.html",
            {
                "product": product,
                "next_product_code": None,
                "error": f"Product code '{product_code}' already exists.",
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/products/", status_code=status.HTTP_303_SEE_OTHER)
