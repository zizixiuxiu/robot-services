"""Product matching utilities for AI order entry."""

from __future__ import annotations

import difflib
from dataclasses import dataclass
from decimal import Decimal
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from sqlalchemy.orm import Session

from app.models.product import Product
from app.schemas.product import ProductCreate
from app.services.products import ProductService


@dataclass
class MatchResult:
    product: Product | None = None
    product_id: int | None = None
    score: float = 0.0
    created_new: bool = False
    warning: str = ""


def _normalize(text: str | None) -> str:
    return (text or "").strip().lower()


def _token_similarity(a: str, b: str) -> float:
    """Return a similarity score between two strings (0.0 - 1.0)."""
    a = _normalize(a)
    b = _normalize(b)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def find_best_product_match(
    db: Session,
    name: str,
    spec: str,
    min_score: float = 0.6,
) -> MatchResult:
    """Find the best matching product by name + spec using fuzzy matching.

    Strategy:
      1. Exact match by name + spec.
      2. Fuzzy match across all products, scoring name and spec independently.
      3. Return the best candidate if combined score >= min_score.
    """
    products = ProductService.list_products(db)
    if not products:
        return MatchResult(warning="产品资料为空，无法匹配。")

    # 1. exact match
    exact = ProductService.get_product_by_name_and_spec(db, name, spec)
    if exact:
        return MatchResult(product=exact, product_id=exact.id, score=1.0)

    # 2. fuzzy match
    best: Product | None = None
    best_score = 0.0
    for product in products:
        name_score = _token_similarity(name, product.product_name)
        spec_score = _token_similarity(spec, product.spec)
        # Weight name a bit more than spec because OCR names are usually more reliable than specs.
        combined = name_score * 0.6 + spec_score * 0.4
        if combined > best_score:
            best_score = combined
            best = product

    if best and best_score >= min_score:
        return MatchResult(product=best, product_id=best.id, score=round(best_score, 3))

    return MatchResult(warning=f"未找到匹配产品：{name} / {spec}")


def ensure_product_for_item(
    db: Session,
    name: str,
    spec: str,
    unit: str = "张",
    unit_price: Decimal = Decimal("0"),
    substrate_category: str = "",
    surface_type: str = "",
    remark: str = "",
    min_score: float = 0.75,
) -> MatchResult:
    """Return a matching product, or create a new one if no good match exists."""
    match = find_best_product_match(db, name, spec, min_score=min_score)
    if match.product:
        return match

    # No good match: create a new product.
    try:
        product = ProductService.create_product(
            db,
            ProductCreate(
                product_code="",
                product_name=name,
                spec=spec,
                unit=unit,
                unit_price=unit_price,
                substrate_category=substrate_category,
                surface_type=surface_type,
                remark=remark or None,
            ),
        )
        return MatchResult(
            product=product,
            product_id=product.id,
            score=0.0,
            created_new=True,
        )
    except ValueError as exc:
        return MatchResult(warning=str(exc))
