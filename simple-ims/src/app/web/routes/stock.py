from fastapi import APIRouter, Request, Depends, Form, status
from fastapi.responses import RedirectResponse
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from app.db import get_db
from app.services.stock import StockService
from app.services.products import ProductService
from app.schemas.stock import StockInCreate, StockOutCreate

router = APIRouter(prefix="/stock")
templates = create_templates()


@router.get("/")
def list_stock(request: Request, db: Session = Depends(get_db)):
    balances = StockService.list_balances(db)
    products = ProductService.list_products(db)
    product_map = {p.id: p for p in products}
    return templates.TemplateResponse(
        request, "stock/list.html", {"balances": balances, "product_map": product_map}
    )


@router.get("/movements")
def list_movements(request: Request, db: Session = Depends(get_db)):
    movements = StockService.list_movements(db)
    products = ProductService.list_products(db)
    product_map = {p.id: p for p in products}
    return templates.TemplateResponse(
        request,
        "stock/movements.html",
        {"movements": movements, "product_map": product_map},
    )


@router.get("/in")
def stock_in_form(request: Request, db: Session = Depends(get_db)):
    products = ProductService.list_products(db)
    return templates.TemplateResponse(
        request, "stock/in_form.html", {"products": products, "error": None}
    )


@router.post("/in")
def stock_in(
    request: Request,
    product_id: int = Form(),
    qty: int = Form(),
    remark: str = Form(default=""),
    db: Session = Depends(get_db),
):
    if qty <= 0:
        products = ProductService.list_products(db)
        return templates.TemplateResponse(
            request,
            "stock/in_form.html",
            {
                "products": products,
                "error": "Quantity must be positive.",
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    data = StockInCreate(product_id=product_id, qty=qty, remark=remark or None)
    StockService.stock_in(db, data)
    return RedirectResponse(url="/stock/", status_code=status.HTTP_303_SEE_OTHER)


@router.get("/out")
def stock_out_form(request: Request, db: Session = Depends(get_db)):
    products = ProductService.list_products(db)
    return templates.TemplateResponse(
        request, "stock/out_form.html", {"products": products, "error": None}
    )


@router.post("/out")
def stock_out(
    request: Request,
    product_id: int = Form(),
    qty: int = Form(),
    remark: str = Form(default=""),
    db: Session = Depends(get_db),
):
    if qty <= 0:
        products = ProductService.list_products(db)
        return templates.TemplateResponse(
            request,
            "stock/out_form.html",
            {
                "products": products,
                "error": "Quantity must be positive.",
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    data = StockOutCreate(product_id=product_id, qty=qty, remark=remark or None)
    try:
        StockService.stock_out(db, data)
    except ValueError as e:
        products = ProductService.list_products(db)
        return templates.TemplateResponse(
            request,
            "stock/out_form.html",
            {
                "products": products,
                "error": str(e),
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/stock/", status_code=status.HTTP_303_SEE_OTHER)
