#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convert an original quotation workbook into a workshop-order workbook.

Rules implemented:
1. Tiepi orders use the full workshop conversion rules.
2. Hunyou orders only clear wood-box packaging and color-adjustment fee rows.
3. Legacy .xls files are converted through LibreOffice before processing so formatting is preserved.

Usage:
    python make_workshop_order.py input.xlsx
    python make_workshop_order.py input.xlsx output.xlsx
"""

from __future__ import annotations

import argparse
import ast
import datetime
import operator
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import xml.etree.ElementTree as ET
import zipfile
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel as to_excel_serial
from openpyxl.worksheet.cell_range import CellRange


DEFAULT_DISCOUNT = 0.8
BUSINESS_SCAN_MAX_COL = 32
LONG_NOTE_MIN_LENGTH = 30
DEFAULT_COLUMN_WIDTH = 8.43
SOFFICE_PROFILE_DIR = Path(
    os.getenv("WORKSHOP_SOFFICE_PROFILE_DIR", str(Path(tempfile.gettempdir()) / "workshop-order-libreoffice"))
)
SOFFICE_TIMEOUT_SECONDS = int(os.getenv("WORKSHOP_SOFFICE_TIMEOUT_SECONDS", "180"))
_SOFFICE_LOCK = threading.RLock()
_SOFFICE_WARMED = False
# openpyxl 保存时会丢掉公式的缓存结果（<v>），只留下公式本身。
# Excel/WPS 打开会重算所以看不出来，但飞书网页版/手机端预览不重算，
# 那些列就整列空白（员工反馈「文件没内容」就是这么来的）。
CACHED_VALUES_ENABLED = os.getenv("WORKSHOP_EMBED_CACHED_VALUES", "1").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}

KW_PRODUCT = "\u4ea7\u54c1"
KW_UNIT_PRICE = "\u5355\u4ef7"
KW_YUAN = "\u5143"
KW_SUBTOTAL = "\u5c0f\u8ba1"
KW_WORKPOINT = "\u5de5\u5206"
KW_WOOD_BOX = "\u6728\u7bb1\u5305\u88c5"
KW_NON_STANDARD_COLOR_FEE = "\u975e\u6807\u8272\u8d39"
KW_TINTING_FEE = "\u8c03\u8272\u8d39"
KW_HUNYOU = "\u6df7\u6cb9"
KW_COLOR = "\u989c\u8272"
KW_MATERIAL = "\u6750\u8d28"
KW_TOTAL = "\u5408\u8ba1"
KW_BOARD = "\u677f\u6750"
KW_HARDWARE = "\u4e94\u91d1"
KW_PAYMENT_TOTAL = "\u5408\u8ba1\u91d1\u989d"
KW_NEED_PAY = "\u9700\u652f\u4ed8"
KW_NEED_PAY_SHORT = "\u9700\u4ed8"
KW_DISCOUNTED = "\u6298\u540e"
KW_ORIGINAL_ORDER = "\u539f\u5355\u53f7"
KW_COLOR_DIFF = "\u8272\u5dee"
KW_AFTER_SALES = "\u552e\u540e"
KW_AGREE_BOARD = "\u540c\u610f\u677f\u6750"
KW_CRAFT = "\u5de5\u827a"
WORKPOINT_DECIMAL_FORMAT = "0.00;[Red]0.00"

# Wood-box row detection must only look at short label text. Long notes
# (e.g. tu-wen-shuo-ming mentioning wood-box packaging) must not trigger
# row clearing, otherwise page-total formulas on the same row get wiped.
WOOD_BOX_LABEL_MAX_LENGTH = 12

DROP_NOTE_KEYWORDS = (
    KW_ORIGINAL_ORDER,
)
PROTECTED_NOTE_KEYWORDS = (
    KW_AFTER_SALES,
    KW_CRAFT,
    "\u4ee5\u5b9e\u7269\u4e3a\u51c6",
)

PAGE_TEXT_RE = re.compile(r"\u6b64\u5355\u5171\s*(\d+)\s*\u9875")
PAGE_COUNT_RE = re.compile(r"\u5171\u8ba1\s*(\d+)\s*\u9875")
# 订单号可能带「加急 / 急」前缀（如「加急S2610-6059-3」）。前缀必须算进号码里，
# 否则 find_existing_order_number 判定失败、退回文件名编号，加急标记和 -N 尾号都会丢。
URGENT_PREFIX = r"(?:\u52a0\u6025|\u6025)?\s*"
ORDER_NUMBER_RE = re.compile(r"^\s*" + URGENT_PREFIX + r"([A-Za-z]\d{4}(?:-\d+)+)")
FULL_ORDER_NUMBER_RE = re.compile(r"^\s*" + URGENT_PREFIX + r"[A-Za-z]\d{4}(?:-\d+)+$")

# Payment-related keywords that should be stripped from bottom notes.
PAYMENT_AMOUNT_KEYWORDS = (
    "\u5408\u8ba1\u91d1\u989d",
    "85\u6298\u4f18\u60e0",
    "8\u6298\u4f18\u60e0",
    "\u9700\u652f\u4ed8",
    "\u9700\u4ed8",
    "\u5e94\u652f\u4ed8",
    "\u5e94\u4ed8",
    "\u4f18\u60e0\u91d1\u989d",
    "\u6298\u540e\u91d1\u989d",
)

PAYMENT_AMOUNT_PATTERNS = (
    re.compile(r"\u6b64\u5355[^，。,；;!！]*(?:\u91d1\u989d|\u91d1\u603b)"),
    re.compile(r"\d+(?:\.\d+)?\s*\u5143"),
    re.compile(r"\d+(?:\.\d+)?\s*\u6298"),
    re.compile(r"\d+(?:\.\d+)?\s*\u6298\s*\u4f18\u60e0"),
    re.compile(r"\u4f18\u60e0[^，。,；;!！]*\d+(?:\.\d+)?\s*\u5143"),
    re.compile(r"\u6298\u540e[^，。,；;!！]*\d+(?:\.\d+)?\s*\u5143"),
    re.compile(r"\u5408\u8ba1\u91d1\u989d[^，。,；;!！]*\d+(?:\.\d+)?\s*\u5143"),
    re.compile(r"\u9700\u652f?\u4ed8[^，。,；;!！]*\d+(?:\.\d+)?\s*\u5143"),
)

PAYMENT_AMOUNT_EMPTY_REMAINDERS = {
    "\u94f6\u955c",
    "\u6837\u54c1\u6253",
    "\u6b64\u5355\u91d1\u989d\u603b\u540c\u610f",
    "\u6b64\u5355\u91d1\u603b\u540c\u610f",
}

PAYMENT_NOTE_TRIM_CHARS = ",\uff0c\uff1b;:\uff1a "
PAYMENT_SEGMENT_SEPARATOR_RE = re.compile(r"(?:[\r\n]+|[ \t\u3000]{4,})")
DROP_NOTE_SEGMENT_KEYWORDS: tuple[str, ...] = ()


@dataclass(frozen=True)
class TextCleanupResult:
    value: str | None
    changed: bool


@dataclass
class TransformStats:
    order_type: str
    discounted_prices: int = 0
    cleared_wood_boxes: int = 0
    cleared_non_standard_color_fees: int = 0
    updated_page_notes: int = 0
    summary_constants_removed: int = 0
    cleared_bottom_notes: int = 0
    filled_workpoints: int = 0
    skipped_prices: int = 0
    deleted_blank_rows: int = 0
    frozen_external_links: int = 0
    moved_hardware_sheets: int = 0
    updated_order_numbers: int = 0
    cached_values_embedded: int = 0
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, int | str | list[str]]:
        return {
            "order_type": self.order_type,
            "discounted_prices": self.discounted_prices,
            "cleared_wood_boxes": self.cleared_wood_boxes,
            "cleared_non_standard_color_fees": self.cleared_non_standard_color_fees,
            "updated_page_notes": self.updated_page_notes,
            "summary_constants_removed": self.summary_constants_removed,
            "cleared_bottom_notes": self.cleared_bottom_notes,
            "filled_workpoints": self.filled_workpoints,
            "skipped_prices": self.skipped_prices,
            "deleted_blank_rows": self.deleted_blank_rows,
            "frozen_external_links": self.frozen_external_links,
            "moved_hardware_sheets": self.moved_hardware_sheets,
            "updated_order_numbers": self.updated_order_numbers,
            "cached_values_embedded": self.cached_values_embedded,
            "warnings": self.warnings,
        }


@dataclass(frozen=True)
class OrderContext:
    input_path: Path
    discount: float
    requested_order_type: str
    resolved_order_type: str
    page_count: str | None
    inferred_order_number: str | None


@dataclass(frozen=True)
class DetailSheetContext:
    wb: Any
    ws: Any
    cached_ws: Any | None
    header_row: int
    product_col: int
    price_col: int
    subtotal_col: int
    workpoint_col: int | None
    color_col: int | None


def payment_amount_info_index(value: str) -> int | None:
    indexes = [idx for kw in PAYMENT_AMOUNT_KEYWORDS if (idx := value.find(kw)) != -1]
    indexes.extend(match.start() for pattern in PAYMENT_AMOUNT_PATTERNS if (match := pattern.search(value)))
    return min(indexes) if indexes else None


def clean_payment_segment(value: str) -> str | None:
    min_idx = payment_amount_info_index(value)
    if min_idx is None:
        return value

    # Keep genuine page/after-sales text, but remove trailing payment details.
    prefix = value[:min_idx].rstrip(PAYMENT_NOTE_TRIM_CHARS)
    keep_prefix = (
        any(keyword in value[:min_idx] for keyword in PROTECTED_NOTE_KEYWORDS)
        or PAGE_TEXT_RE.search(prefix)
        or PAGE_COUNT_RE.search(prefix)
    )
    if keep_prefix and prefix and prefix not in PAYMENT_AMOUNT_EMPTY_REMAINDERS:
        # Discount approval text can be glued to an after-sales sentence, e.g.
        # "...不进入公司售后！此单某总同意板材95折优惠...". Keep only the
        # completed after-sales/page sentence before that price fragment.
        delimiter_positions = [prefix.rfind(ch) for ch in "\u3002\uff01!?\uff1f\uff1b;"]
        last_delimiter = max(delimiter_positions)
        if last_delimiter != -1:
            completed_prefix = prefix[: last_delimiter + 1].rstrip(PAYMENT_NOTE_TRIM_CHARS)
            dangling_fragment = prefix[last_delimiter + 1 :].strip()
            dangling_has_page_count = PAGE_TEXT_RE.search(dangling_fragment) or PAGE_COUNT_RE.search(dangling_fragment)
            if completed_prefix and dangling_fragment and KW_AFTER_SALES not in dangling_fragment and not dangling_has_page_count:
                return completed_prefix
        return prefix

    # If the segment is only meaningful because of the discount/amount, drop it entirely.
    return None


def clean_payment_amount_text(value: str) -> TextCleanupResult:
    """Remove payment/discount amounts from a note."""
    min_idx = payment_amount_info_index(value)
    if min_idx is None:
        return TextCleanupResult(value, False)

    segments = [segment for segment in PAYMENT_SEGMENT_SEPARATOR_RE.split(value) if segment.strip()]
    if len(segments) > 1:
        cleaned_segments = []
        changed = False
        for segment in segments:
            cleaned_segment = clean_payment_segment(segment.strip())
            if cleaned_segment is None:
                changed = True
                continue
            changed = changed or cleaned_segment != segment.strip()
            cleaned_segments.append(cleaned_segment)

        cleaned = "\n".join(cleaned_segments).strip()
        if not cleaned:
            return TextCleanupResult(None, True)
        return TextCleanupResult(cleaned, changed or cleaned != value)

    cleaned = clean_payment_segment(value)
    if not cleaned or cleaned in PAYMENT_AMOUNT_EMPTY_REMAINDERS:
        return TextCleanupResult(None, True)
    return TextCleanupResult(cleaned, cleaned != value)


def remove_payment_amount_info(value: str) -> str:
    """Strip payment amount suffixes from bottom notes."""
    result = clean_payment_amount_text(value)
    return result.value or ""


def should_drop_note_segment(segment: str) -> bool:
    if not any(keyword in segment for keyword in DROP_NOTE_SEGMENT_KEYWORDS):
        return False
    return "\u6b64\u5355" in segment or len(segment.strip()) >= LONG_NOTE_MIN_LENGTH


def clean_non_workshop_note_text(value: str) -> TextCleanupResult:
    """Drop after-sales/process note segments from workshop output notes."""
    segments = [segment for segment in PAYMENT_SEGMENT_SEPARATOR_RE.split(value) if segment.strip()]
    if not segments:
        return TextCleanupResult(value, False)

    cleaned_segments = []
    changed = False
    for segment in segments:
        stripped = segment.strip()
        if should_drop_note_segment(stripped):
            changed = True
            continue
        cleaned_segments.append(stripped)

    if not changed:
        return TextCleanupResult(value, False)
    cleaned = "\n".join(cleaned_segments).strip()
    return TextCleanupResult(cleaned or None, True)


ORDER_TYPES = {"auto", "tiepi", "hunyou"}
COLOR_FEE_KEYWORDS = (
    KW_NON_STANDARD_COLOR_FEE,
    KW_TINTING_FEE,
)


def text(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def safe_eval_arithmetic(expr: str) -> float | None:
    """Evaluate simple numeric formulas such as '=794+120+150+80'."""
    if not expr or not expr.startswith("="):
        return None

    source = expr[1:].strip()
    if not re.fullmatch(r"[0-9+\-*/().\s]+", source):
        return None

    operators = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.USub: operator.neg,
        ast.UAdd: operator.pos,
    }

    def visit(node: ast.AST) -> float:
        if isinstance(node, ast.Expression):
            return visit(node.body)
        if isinstance(node, ast.Constant) and is_number(node.value):
            return float(node.value)
        if isinstance(node, ast.UnaryOp) and type(node.op) in operators:
            return operators[type(node.op)](visit(node.operand))
        if isinstance(node, ast.BinOp) and type(node.op) in operators:
            return operators[type(node.op)](visit(node.left), visit(node.right))
        raise ValueError("unsupported formula")

    try:
        return visit(ast.parse(source, mode="eval"))
    except Exception:
        return None


def find_header_columns(ws: Any) -> tuple[int, int, int, int, int | None] | None:
    """Return header_row, product_col, price_col, subtotal_col, workpoint_col."""
    max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
    for row in ws.iter_rows(max_col=max_col):
        product_col = price_col = subtotal_col = workpoint_col = None
        for cell in row:
            value = text(cell.value).replace("\n", "")
            # Accept both "单价" / "单价（元）" and "小计" / "小计（元）" forms.
            if KW_PRODUCT in value:
                product_col = cell.column
            elif KW_UNIT_PRICE in value:
                price_col = cell.column
            elif KW_SUBTOTAL in value:
                subtotal_col = cell.column
            elif KW_WORKPOINT in value:
                workpoint_col = cell.column
        if product_col and price_col and subtotal_col:
            return row[0].row, product_col, price_col, subtotal_col, workpoint_col
    return None


def get_price_value(formula_cell: Any, cached_cell: Any | None = None) -> float | None:
    raw = formula_cell.value
    if is_number(raw):
        return float(raw)
    if isinstance(raw, str):
        price = safe_eval_arithmetic(raw)
        if price is not None:
            return price
    if cached_cell is not None and is_number(cached_cell.value):
        return float(cached_cell.value)
    return None


def remove_numeric_formula_terms(formula: str) -> str:
    """Remove standalone numeric additions from a board-summary formula, e.g. '+160'."""
    if not formula or not formula.startswith("="):
        return formula
    result = re.sub(r"\+\s*\d+(?:\.\d+)?(?=$|[+\-])", "", formula)
    result = re.sub(r"-\s*\d+(?:\.\d+)?(?=$|[+\-])", "", result)
    return result


def is_color_fee_product(product: str) -> bool:
    return any(keyword in product for keyword in COLOR_FEE_KEYWORDS)


def is_hunyou_color(color: str) -> bool:
    return re.match(r"^[\s（(【\[]*\s*混油", color) is not None


def detect_order_type(wb: Any, requested: str) -> str:
    if requested != "auto":
        return requested

    for ws in wb.worksheets:
        # Only treat sheets with a full production-detail header as production
        # lists. BOM/summary sheets may also contain "产品名称" and "颜色"
        # columns but lack a "单价（元）" column, so they must not influence
        # order-type detection.
        headers = find_header_columns(ws)
        if not headers:
            continue
        header_row, product_col, price_col, subtotal_col, workpoint_col = headers

        color_col = None
        for cell in ws[header_row]:
            value = text(cell.value).replace("\n", "")
            if KW_COLOR in value:
                color_col = cell.column
                break
        if not color_col:
            continue

        current_product = ""
        current_color = ""
        for row_idx in range(header_row + 1, ws.max_row + 1):
            raw_product = text(ws.cell(row_idx, product_col).value) if product_col else ""
            if raw_product:
                current_product = raw_product
            product = current_product

            if KW_WOOD_BOX in product or product == KW_TOTAL:
                break

            raw_color = text(ws.cell(row_idx, color_col).value)
            if raw_color:
                current_color = raw_color
            color = current_color

            if is_hunyou_color(color):
                return "hunyou"
    return "tiepi"


def find_color_column(ws: Any) -> tuple[int, int, int | None] | None:
    """Return header_row, color_col, product_col for production detail sheets."""
    max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
    for row in ws.iter_rows(max_col=max_col):
        color_col = None
        product_col = None
        for cell in row:
            value = text(cell.value).replace("\n", "")
            if KW_COLOR in value:
                color_col = cell.column
            elif KW_PRODUCT in value:
                product_col = cell.column
        if color_col and product_col:
            return row[0].row, color_col, product_col
    return None


def extract_page_count(wb: Any) -> str | None:
    for ws in wb.worksheets:
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for row in ws.iter_rows(max_col=max_col):
            for cell in row:
                value = text(cell.value)
                match = PAGE_TEXT_RE.search(value) or PAGE_COUNT_RE.search(value)
                if match:
                    return match.group(1)
    return None


def infer_order_number(input_path: Path) -> str | None:
    match = ORDER_NUMBER_RE.match(input_path.stem)
    return match.group(1) if match else None


def is_full_order_number(value: Any) -> bool:
    return isinstance(value, str) and FULL_ORDER_NUMBER_RE.fullmatch(value.strip()) is not None


def find_existing_order_number(wb: Any) -> str | None:
    """Return a concrete order number already present in the workbook header."""
    for ws in wb.worksheets:
        if not str(ws.title).isdigit():
            continue
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for row in ws.iter_rows(max_col=max_col):
            for cell in row:
                if text(cell.value) != "\u8ba2\u5355\u7f16\u53f7":
                    continue
                for offset in (1, 2):
                    value = ws.cell(cell.row, cell.column + offset).value
                    if is_full_order_number(value):
                        return str(value).strip()
    return None


def set_plain_cell(ws: Any, row: int, col: int, value: Any) -> bool:
    if row < 1 or col < 1 or row > ws.max_row or col > ws.max_column:
        return False
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell) or cell.value == value:
        return False
    cell.value = value
    return True


def should_update_suffix_cell(ws: Any, row: int, col: int) -> bool:
    if row < 1 or col < 1 or row > ws.max_row or col > ws.max_column:
        return False
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        return False
    return cell.value not in (None, "")


def normalize_order_numbers(wb: Any, order_number: str | None) -> int:
    if not order_number:
        return 0

    changed = 0
    for ws in wb.worksheets:
        if not str(ws.title).isdigit():
            continue
        suffix = -int(ws.title)
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for row in ws.iter_rows(max_col=max_col):
            for cell in row:
                value = text(cell.value)
                if value == "\u8ba2\u5355\u7f16\u53f7":
                    changed += int(set_plain_cell(ws, cell.row, cell.column + 1, order_number))
                    if should_update_suffix_cell(ws, cell.row, cell.column + 2):
                        changed += int(set_plain_cell(ws, cell.row, cell.column + 2, suffix))
                elif value == "\u751f\u4ea7\u5355\u53f7" and cell.column >= 14:
                    changed += int(set_plain_cell(ws, cell.row, cell.column + 2, order_number))
                    if should_update_suffix_cell(ws, cell.row, cell.column + 4):
                        changed += int(set_plain_cell(ws, cell.row, cell.column + 4, suffix))
    return changed


def translate_formula(formula: str, origin: str, target: str) -> str | None:
    try:
        return Translator(formula, origin=origin).translate_formula(target)
    except Exception:
        return None


CELL_REF_RE = re.compile(
    r"(?P<sheet>(?:'(?P<quoted_sheet>(?:[^']|'')+)'|(?P<plain_sheet>[^'!:+\-*/^&=<>(),\s]+))!)?"
    r"(?P<col_abs>\$?)(?P<col>[A-Z]{1,3})(?P<row_abs>\$?)(?P<row>\d+)"
)


def formula_sheet_name(match: re.Match[str]) -> str | None:
    if match.group("quoted_sheet") is not None:
        return match.group("quoted_sheet").replace("''", "'")
    if match.group("plain_sheet") is not None:
        return match.group("plain_sheet")
    return None


def adjust_formula_after_row_delete(
    formula: str,
    formula_sheet: str,
    deleted_sheet: str,
    deleted_row: int,
    amount: int = 1,
) -> str:
    """Update A1-style row references after deleting rows from one sheet."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    def replace(match: re.Match[str]) -> str:
        ref_sheet = formula_sheet_name(match) or formula_sheet
        if ref_sheet != deleted_sheet:
            return match.group(0)

        row = int(match.group("row"))
        if row > deleted_row:
            row -= amount
        elif row == deleted_row and match.start() > 0 and formula[match.start() - 1] == ":":
            row = max(1, row - amount)
        else:
            return match.group(0)

        return (
            f"{match.group('sheet') or ''}"
            f"{match.group('col_abs')}{match.group('col')}"
            f"{match.group('row_abs')}{row}"
        )

    return CELL_REF_RE.sub(replace, formula)


def adjust_workbook_formulas_after_row_delete(wb: Any, deleted_ws: Any, deleted_row: int, amount: int = 1) -> int:
    changed = 0
    for ws in wb.worksheets:
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for row in ws.iter_rows(max_col=max_col):
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("="):
                    continue
                adjusted = adjust_formula_after_row_delete(value, ws.title, deleted_ws.title, deleted_row, amount)
                if adjusted != value:
                    cell.value = adjusted
                    changed += 1
    return changed


def _merged_range_address(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    return (
        f"{get_column_letter(min_col)}{min_row}:"
        f"{get_column_letter(max_col)}{max_row}"
    )


def prepare_merged_cells_for_row_delete(ws: Any, deleted_row: int, amount: int = 1) -> list[str]:
    shifted_ranges: list[str] = []
    for merged_range in list(ws.merged_cells.ranges):
        cell_range = CellRange(str(merged_range))
        if cell_range.min_row <= deleted_row <= cell_range.max_row:
            ws.unmerge_cells(str(merged_range))
            # Keep the merge, just shrink it. Dropping it leaves quote color/style
            # blocks as a single top-left cell (e.g. only N8 shows "混油附色").
            new_min = deleted_row if cell_range.min_row == deleted_row else cell_range.min_row
            new_max = cell_range.max_row - amount
            if new_max >= new_min and (new_max > new_min or cell_range.max_col > cell_range.min_col):
                shifted_ranges.append(
                    _merged_range_address(cell_range.min_col, new_min, cell_range.max_col, new_max)
                )
        elif cell_range.min_row > deleted_row:
            ws.unmerge_cells(str(merged_range))
            cell_range.shift(row_shift=-amount)
            shifted_ranges.append(str(cell_range))
    return shifted_ranges


def row_has_size_or_qty(ws: Any, row_idx: int) -> bool:
    """True when a continuation row still has height/width/thickness/qty."""
    for col_idx in range(5, 9):  # E-H in the quote template
        value = ws.cell(row_idx, col_idx).value
        if isinstance(value, bool) or value in (None, ""):
            continue
        if isinstance(value, (int, float)):
            return True
        text_value = text(value)
        if text_value and text_value[0].isdigit():
            return True
    return False


def last_named_product_row(ws: Any, header_row: int, product_col: int) -> int | None:
    last: int | None = None
    for row_idx in range(header_row + 1, ws.max_row + 1):
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        row_values = [text(ws.cell(row_idx, col_idx).value) for col_idx in range(1, max_col + 1)]
        if any(value == KW_TOTAL or (KW_WOOD_BOX in value and len(value) <= WOOD_BOX_LABEL_MAX_LENGTH) for value in row_values):
            break
        product = text(ws.cell(row_idx, product_col).value)
        if product and product != KW_TOTAL:
            last = row_idx
            merged = merged_range_for_cell(ws, row_idx, product_col)
            if merged is not None:
                last = max(last, merged.max_row)
        elif last is not None and row_idx > last and row_has_size_or_qty(ws, row_idx):
            last = row_idx
    return last


def color_block_end_row(
    ws: Any,
    header_row: int,
    product_col: int,
    start_row: int,
    color_min_col: int,
) -> int | None:
    """Last row the color block must cover.

    Product names, vertically merged names (条子), size continuation rows,
    an already-correct color merge, and a spanning 材质 merge are all
    treated as the same product block so restore never shrinks color.
    """
    end_row = last_named_product_row(ws, header_row, product_col)
    if end_row is None:
        return None

    existing = merged_range_for_cell(ws, start_row, color_min_col)
    if existing is not None and existing.min_col == color_min_col:
        end_row = max(end_row, existing.max_row)

    for cell in ws[header_row]:
        if KW_MATERIAL not in text(cell.value).replace("\n", ""):
            continue
        material_merge = merged_range_for_cell(ws, start_row, cell.column)
        if material_merge is not None and material_merge.max_row >= start_row:
            end_row = max(end_row, material_merge.max_row)
        break
    return end_row


def color_block_columns(ws: Any, header_row: int, color_col: int) -> tuple[int, int]:
    merged = merged_range_for_cell(ws, header_row, color_col)
    if merged is not None:
        return merged.min_col, merged.max_col
    next_header = text(ws.cell(header_row, color_col + 1).value)
    if not next_header:
        return color_col, color_col + 1
    return color_col, color_col


def restore_color_block_merge(ws: Any, header_row: int, product_col: int, color_col: int | None) -> bool:
    """Rebuild the quote-template color block covering all product rows.

    Template stores one color in the first data cell and merges N:O down to the
    last product row. LibreOffice conversion or blank-row deletion can drop that
    merge, so later rows look empty and the first cell displays truncated text.
    """
    if not color_col:
        return False

    start_row = header_row + 1
    min_col, max_col = color_block_columns(ws, header_row, color_col)
    end_row = color_block_end_row(ws, header_row, product_col, start_row, min_col)
    if end_row is None or end_row < start_row:
        return False

    # If the color column already carries more than one distinct value inside
    # the product block (e.g. accent colors / 叉色), keep the original
    # segmentation untouched. Rebuilding one big block would wipe those colors.
    distinct_color_values: set[Any] = set()
    color_value: Any | None = None
    color_row: int | None = None
    for row_idx in range(start_row, end_row + 1):
        value = ws.cell(row_idx, min_col).value
        if value not in (None, ""):
            distinct_color_values.add(value)
            if color_value is None:
                color_value = value
                color_row = row_idx
    if len(distinct_color_values) > 1:
        return False
    if color_value in (None, "") or color_row is None:
        return False

    existing = merged_range_for_cell(ws, start_row, min_col)
    if (
        existing is not None
        and existing.min_row == start_row
        and existing.max_row == end_row
        and existing.min_col == min_col
        and existing.max_col == max_col
    ):
        return False

    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.max_col < min_col or merged_range.min_col > max_col:
            continue
        if merged_range.max_row < start_row or merged_range.min_row > end_row:
            continue
        ws.unmerge_cells(str(merged_range))

    if color_row != start_row:
        ws.cell(start_row, min_col).value = color_value
        ws.cell(color_row, min_col).value = None

    ws.merge_cells(_merged_range_address(min_col, start_row, max_col, end_row))
    top_left = ws.cell(start_row, min_col)
    top_left.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return True


def ensure_workpoint_decimal_format(cell: Any) -> bool:
    """工分按实际值显示两位小数，避免整数格式把 32.64 显示成 33。"""
    if isinstance(cell, MergedCell):
        return False
    fmt = str(cell.number_format or "")
    if "0.00" in fmt:
        return False
    cell.number_format = WORKPOINT_DECIMAL_FORMAT
    return True


def fill_workpoint_formulas(
    ws: Any,
    header_row: int,
    product_col: int,
    subtotal_col: int,
    workpoint_col: int | None,
) -> int:
    if not workpoint_col:
        return 0

    filled = 0
    data_start = header_row + 1
    last_product_row = header_row
    known_formula_cell = None
    workpoint_letter = get_column_letter(workpoint_col)
    subtotal_letter = get_column_letter(subtotal_col)
    current_product = ""
    # 小计纵向合并块（门套整组合并单价/小计）内已处理过的行，避免重复清零
    zeroed_merged_rows: set[int] = set()

    def normalized_formula(value: Any) -> str:
        if not isinstance(value, str):
            return ""
        return value.replace(" ", "").replace("=+", "=").upper()

    def vertical_subtotal_merge(row_idx: int) -> CellRange | None:
        """小计列纵向跨多行的合并（整组门洞/侧套/门扇共用一个小计）。"""
        merged = merged_range_for_cell(ws, row_idx, subtotal_col)
        if merged is not None and merged.max_row > merged.min_row:
            return merged
        return None

    def zero_workpoints_in_merge(merge: CellRange) -> int:
        """合并小计块内保留一份工分 =小计/100，其余置 0，避免多行都写 =L顶/100 导致工分翻倍。"""
        changed = 0
        expected_formula = f"={subtotal_letter}{merge.min_row}/100"

        # 先找到应该写工分公式的目标单元格：优先同列、首行开始的真实 Cell。
        # 如果首行是 MergedCell，但其合并区域左上角就在工分列且在本小计合并块内，
        # 则写到那个左上角单元格，保证不会因为首行被合并而整组漏算。
        target_cell: Any | None = None
        target_row: int | None = None
        for row_idx in range(merge.min_row, merge.max_row + 1):
            zeroed_merged_rows.add(row_idx)
            cell = ws.cell(row_idx, workpoint_col)
            if isinstance(cell, MergedCell):
                cell_merge = merged_range_for_cell(ws, row_idx, workpoint_col)
                if (
                    cell_merge is not None
                    and cell_merge.min_col == workpoint_col
                    and merge.min_row <= cell_merge.min_row <= merge.max_row
                ):
                    tl_cell = ws.cell(cell_merge.min_row, cell_merge.min_col)
                    if not isinstance(tl_cell, MergedCell):
                        if target_cell is None or cell_merge.min_row < target_row:
                            target_cell = tl_cell
                            target_row = cell_merge.min_row
                continue
            if target_cell is None or row_idx < target_row:
                target_cell = cell
                target_row = row_idx

        if target_cell is None:
            return 0

        if isinstance(target_cell.value, str) and target_cell.value.startswith("="):
            if normalized_formula(target_cell.value) != normalized_formula(expected_formula):
                target_cell.value = expected_formula
                changed += 1
        elif target_cell.value not in (expected_formula,):
            target_cell.value = expected_formula
            changed += 1
        apply_workpoint_style(target_cell)
        if ensure_workpoint_decimal_format(target_cell):
            changed += 1

        for row_idx in range(merge.min_row, merge.max_row + 1):
            cell = ws.cell(row_idx, workpoint_col)
            if isinstance(cell, MergedCell):
                continue
            if cell.coordinate == target_cell.coordinate:
                continue
            if cell.value != 0:
                cell.value = 0
                changed += 1
        return changed

    def find_workpoint_style_template() -> Any | None:
        scan_product = ""
        for scan_row in range(data_start, ws.max_row + 1):
            max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
            row_labels = [text(ws.cell(scan_row, col).value) for col in range(1, max_col + 1)]
            if any(KW_WOOD_BOX in value or value == KW_TOTAL for value in row_labels):
                break

            if vertical_subtotal_merge(scan_row) is not None:
                continue

            raw_product = text(ws.cell(scan_row, product_col).value)
            subtotal_value = ws.cell(scan_row, subtotal_col).value
            if raw_product:
                scan_product = raw_product
            elif scan_product and subtotal_value not in (None, ""):
                pass
            else:
                scan_product = ""
                continue
            if subtotal_value in (None, ""):
                continue

            candidate = ws.cell(scan_row, workpoint_col)
            expected = f"={subtotal_letter}{scan_row}/100"
            if isinstance(candidate.value, str) and normalized_formula(candidate.value) == normalized_formula(expected):
                return candidate
        return None

    style_template_cell = find_workpoint_style_template()

    def apply_workpoint_style(cell: Any) -> None:
        if style_template_cell is not None and cell.coordinate != style_template_cell.coordinate:
            cell._style = copy(style_template_cell._style)
        ensure_workpoint_decimal_format(cell)

    for row_idx in range(data_start, ws.max_row + 1):
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        row_labels = [text(ws.cell(row_idx, col).value) for col in range(1, max_col + 1)]
        if any(KW_WOOD_BOX in value or value == KW_TOTAL for value in row_labels):
            break

        merge = vertical_subtotal_merge(row_idx)
        if merge is not None:
            last_product_row = max(last_product_row, merge.max_row)
            if row_idx not in zeroed_merged_rows:
                filled += zero_workpoints_in_merge(merge)
            current_product = ""
            continue

        raw_product = text(ws.cell(row_idx, product_col).value)
        subtotal_value = ws.cell(row_idx, subtotal_col).value
        if raw_product:
            current_product = raw_product
        elif current_product and subtotal_value not in (None, ""):
            # continuation of a vertically merged product cell
            pass
        else:
            current_product = ""
            continue
        if subtotal_value in (None, ""):
            continue

        product = current_product
        last_product_row = row_idx
        workpoint_cell = ws.cell(row_idx, workpoint_col)
        # 某些 .xls 经 LibreOffice 转换后，工分列会被跨列合并块（如 M8:O24）覆盖，
        # 此时 MergedCell 不可写。跳过这些行，避免崩溃；工分由左上角承担即可。
        if isinstance(workpoint_cell, MergedCell):
            continue
        expected_formula = f"={subtotal_letter}{row_idx}/100"
        if isinstance(workpoint_cell.value, str) and workpoint_cell.value.startswith("="):
            if normalized_formula(workpoint_cell.value) != normalized_formula(expected_formula):
                workpoint_cell.value = expected_formula
                filled += 1
            apply_workpoint_style(workpoint_cell)
            if ensure_workpoint_decimal_format(workpoint_cell):
                filled += 1
            known_formula_cell = workpoint_cell
            continue
        if workpoint_cell.value not in (None, ""):
            continue

        formula = None
        if known_formula_cell is not None:
            formula = translate_formula(
                known_formula_cell.value,
                known_formula_cell.coordinate,
                workpoint_cell.coordinate,
            )
        if not formula:
            formula = f"={subtotal_letter}{row_idx}/100"

        workpoint_cell.value = formula
        apply_workpoint_style(workpoint_cell)
        filled += 1

    if last_product_row <= header_row:
        return filled

    for row_idx in range(last_product_row + 1, ws.max_row + 1):
        labels = [text(ws.cell(row_idx, col).value) for col in range(1, min(ws.max_column, product_col + 1) + 1)]
        if any(value == KW_TOTAL for value in labels):
            total_cell = ws.cell(row_idx, workpoint_col)
            if total_cell.value in (None, ""):
                total_cell.value = f"=SUM({workpoint_letter}{data_start}:{workpoint_letter}{last_product_row})"
                filled += 1
            break

    return filled


def delete_blank_rows_between_data_and_summary(wb: Any, ws: Any, header_row: int) -> int:
    """Delete blank rows between product data and summary/wood-box rows.

    Merged cells that span the blank row are unmerged first. Formula references
    across the workbook are then shifted with Excel-like row-deletion semantics.
    """
    boundary_row = None
    for row_idx in range(header_row + 1, ws.max_row + 1):
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for col_idx in range(1, max_col + 1):
            val = text(ws.cell(row_idx, col_idx).value)
            if val in (KW_TOTAL, KW_WOOD_BOX):
                boundary_row = row_idx
                break
        if boundary_row:
            break

    if boundary_row is None or boundary_row <= header_row + 1:
        return 0

    deleted = 0
    for row_idx in range(boundary_row - 1, header_row, -1):
        is_blank = True
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row_idx, col_idx)
            if cell.value not in (None, ""):
                is_blank = False
                break
        if not is_blank:
            continue

        shifted_merged_ranges = prepare_merged_cells_for_row_delete(ws, row_idx)

        ws.delete_rows(row_idx)
        for merged_range in shifted_merged_ranges:
            ws.merge_cells(merged_range)
        adjust_workbook_formulas_after_row_delete(wb, ws, row_idx)

        # Keep sequence numbers in column A continuous after the row deletion.
        for r in range(row_idx, ws.max_row + 1):
            seq_cell = ws.cell(r, 1)
            if isinstance(seq_cell.value, int):
                seq_cell.value -= 1

        deleted += 1

    return deleted


def is_hardware_sheet_name(sheet_name: str) -> bool:
    name = str(sheet_name).strip()
    return name.startswith(KW_HARDWARE) or name.endswith(KW_HARDWARE) or KW_HARDWARE in name


def move_hardware_sheets_to_end(wb: Any) -> int:
    moved = 0
    for ws in list(wb.worksheets):
        if is_hardware_sheet_name(ws.title):
            wb._sheets.remove(ws)
            wb._sheets.append(ws)
            moved += 1
    return moved


def build_order_context(input_path: Path, discount: float, requested_order_type: str, wb: Any) -> OrderContext:
    resolved_order_type = detect_order_type(wb, requested_order_type)
    return OrderContext(
        input_path=input_path,
        discount=discount,
        requested_order_type=requested_order_type,
        resolved_order_type=resolved_order_type,
        page_count=extract_page_count(wb),
        inferred_order_number=find_existing_order_number(wb) or infer_order_number(input_path),
    )


def build_detail_sheet_context(wb: Any, cached_wb: Any | None, ws: Any) -> DetailSheetContext | None:
    headers = find_header_columns(ws)
    if not headers:
        return None

    header_row, product_col, price_col, subtotal_col, workpoint_col = headers
    color_header = find_color_column(ws)
    color_col = color_header[1] if color_header else None
    return DetailSheetContext(
        wb=wb,
        ws=ws,
        cached_ws=cached_wb[ws.title] if cached_wb is not None else None,
        header_row=header_row,
        product_col=product_col,
        price_col=price_col,
        subtotal_col=subtotal_col,
        workpoint_col=workpoint_col,
        color_col=color_col,
    )


def workbook_needs_cached_prices(wb: Any) -> bool:
    for ws in wb.worksheets:
        headers = find_header_columns(ws)
        if not headers:
            continue

        header_row, product_col, price_col, subtotal_col, _workpoint_col = headers
        current_product = ""
        for row_idx in range(header_row + 1, ws.max_row + 1):
            raw_product = text(ws.cell(row_idx, product_col).value)
            price_cell = ws.cell(row_idx, price_col)
            subtotal_cell = ws.cell(row_idx, subtotal_col)

            if raw_product:
                current_product = raw_product
            elif current_product and (price_cell.value not in (None, "") or subtotal_cell.value not in (None, "")):
                pass
            else:
                current_product = ""

            product = current_product
            if (
                is_color_fee_product(product)
                or KW_WOOD_BOX in product
                or not product
                or KW_TOTAL in product
                or isinstance(price_cell, MergedCell)
            ):
                continue

            raw_price = price_cell.value
            if isinstance(raw_price, str) and raw_price.startswith("=") and safe_eval_arithmetic(raw_price) is None:
                return True

    return False


def clear_color_fee_row(ctx: DetailSheetContext, row_idx: int) -> bool:
    last_clear_col = max(
        ctx.subtotal_col,
        ctx.workpoint_col or ctx.subtotal_col,
        min(ctx.ws.max_column, 18),
    )
    changed = False
    for col_idx in range(1, last_clear_col + 1):
        row_cell = ctx.ws.cell(row_idx, col_idx)
        if not isinstance(row_cell, MergedCell) and row_cell.value not in (None, ""):
            row_cell.value = None
            changed = True
    return changed


def clear_wood_box_row(ctx: DetailSheetContext, row_idx: int) -> bool:
    changed = False
    for col_idx in (ctx.subtotal_col, ctx.workpoint_col):
        if not col_idx:
            continue
        row_cell = ctx.ws.cell(row_idx, col_idx)
        if not isinstance(row_cell, MergedCell) and row_cell.value not in (None, ""):
            row_cell.value = None
            changed = True
    return changed


def process_detail_rows(ctx: DetailSheetContext, order_context: OrderContext, stats: TransformStats) -> None:
    current_product = ""
    for row_idx in range(ctx.header_row + 1, ctx.ws.max_row + 1):
        max_col = min(ctx.ws.max_column, BUSINESS_SCAN_MAX_COL)
        row_values = [text(ctx.ws.cell(row_idx, col_idx).value) for col_idx in range(1, max_col + 1)]
        if any(KW_WOOD_BOX in value and len(value) <= WOOD_BOX_LABEL_MAX_LENGTH for value in row_values):
            if clear_wood_box_row(ctx, row_idx):
                stats.cleared_wood_boxes += 1
            current_product = ""
            continue

        raw_product = text(ctx.ws.cell(row_idx, ctx.product_col).value)
        price_cell = ctx.ws.cell(row_idx, ctx.price_col)
        subtotal_cell = ctx.ws.cell(row_idx, ctx.subtotal_col)

        if raw_product:
            current_product = raw_product
        elif current_product and (price_cell.value not in (None, "") or subtotal_cell.value not in (None, "")):
            pass
        else:
            current_product = ""

        product = current_product
        if is_color_fee_product(product):
            if clear_color_fee_row(ctx, row_idx):
                stats.cleared_non_standard_color_fees += 1
            continue

        if KW_WOOD_BOX in product and len(product) <= WOOD_BOX_LABEL_MAX_LENGTH:
            if clear_wood_box_row(ctx, row_idx):
                stats.cleared_wood_boxes += 1
            continue

        if not product or KW_TOTAL in product or isinstance(price_cell, MergedCell):
            continue

        # 混油订单不清除行内容，但不对价格进行折扣
        if order_context.resolved_order_type == "hunyou":
            continue

        cached_price_cell = ctx.cached_ws.cell(row_idx, ctx.price_col) if ctx.cached_ws is not None else None
        price = get_price_value(price_cell, cached_price_cell)
        if price is None:
            if price_cell.value not in (None, ""):
                stats.skipped_prices += 1
            continue

        price_cell.value = price * order_context.discount
        stats.discounted_prices += 1


def cleanup_sheet_notes(ws: Any, order_context: OrderContext, stats: TransformStats) -> None:
    max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
    for row in ws.iter_rows(max_col=max_col):
        row_values = [text(cell.value) for cell in row]
        row_has_board_summary = any(KW_BOARD in value for value in row_values)
        for cell in row:
            value = cell.value
            if not isinstance(value, str):
                continue

            if value.startswith("=") and row_has_board_summary:
                new_formula = remove_numeric_formula_terms(value)
                if new_formula != value:
                    cell.value = new_formula
                    stats.summary_constants_removed += 1
                continue

            note_cleanup = clean_non_workshop_note_text(value)
            if note_cleanup.changed:
                cell.value = note_cleanup.value
                stats.updated_page_notes += 1
                if cell.value is None:
                    continue
                value = cell.value

            cleanup = clean_payment_amount_text(value)
            if cleanup.changed:
                cell.value = cleanup.value
                stats.updated_page_notes += 1


def display_text_units(value: str) -> int:
    return sum(2 if ord(ch) > 127 else 1 for ch in value)


def is_long_note_text(value: str) -> bool:
    if len(value.strip()) < LONG_NOTE_MIN_LENGTH:
        return False
    return KW_AFTER_SALES in value or KW_AGREE_BOARD in value or "\u56fe\u6587\u8bf4\u660e" in value


def normalize_note_display_text(value: str) -> str:
    """Use real line breaks instead of long space runs inside merged note cells."""
    return re.sub(r"[ \t\u3000]{4,}", "\n", value.strip())


def merged_range_for_cell(ws: Any, row_idx: int, col_idx: int) -> CellRange | None:
    coordinate = ws.cell(row_idx, col_idx).coordinate
    for merged_range in ws.merged_cells.ranges:
        if coordinate in merged_range:
            return CellRange(str(merged_range))
    return None


def cell_display_width(ws: Any, row_idx: int, col_idx: int) -> float:
    merged_range = merged_range_for_cell(ws, row_idx, col_idx)
    min_col = merged_range.min_col if merged_range else col_idx
    max_col = merged_range.max_col if merged_range else col_idx
    width = 0.0
    for col in range(min_col, max_col + 1):
        letter = get_column_letter(col)
        width += ws.column_dimensions[letter].width or DEFAULT_COLUMN_WIDTH
    return max(width, DEFAULT_COLUMN_WIDTH)


def estimated_wrapped_lines(value: str, width: float) -> int:
    capacity = max(12, int(width * 1.15))
    lines = 0
    for part in value.splitlines() or [""]:
        units = max(1, display_text_units(part))
        lines += max(1, (units + capacity - 1) // capacity)
    return lines


def fit_long_note_rows(ws: Any) -> int:
    adjusted = 0
    max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
    for row in ws.iter_rows(max_col=max_col):
        for cell in row:
            value = cell.value
            if not isinstance(value, str) or not is_long_note_text(value):
                continue

            normalized = normalize_note_display_text(value)
            if normalized != value:
                cell.value = normalized
                value = normalized

            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = alignment.vertical or "center"
            cell.alignment = alignment

            width = cell_display_width(ws, cell.row, cell.column)
            lines = estimated_wrapped_lines(value, width)
            desired_height = min(180, max(42, lines * 24 + 18))

            merged_range = merged_range_for_cell(ws, cell.row, cell.column)
            if merged_range is not None and merged_range.max_row > cell.row:
                # 纵向跨多行的合并单元格：按整个合并区域的合计行高判断，
                # 不够时把差值补到首行，避免把全部高度堆到一行
                spanned_rows = range(cell.row, merged_range.max_row + 1)
                total_height = sum(
                    ws.row_dimensions[r].height or 15 for r in spanned_rows
                )
                if total_height < desired_height:
                    current_height = ws.row_dimensions[cell.row].height or 15
                    ws.row_dimensions[cell.row].height = (
                        current_height + desired_height - total_height
                    )
                    adjusted += 1
                continue

            current_height = ws.row_dimensions[cell.row].height or 15
            if current_height < desired_height:
                ws.row_dimensions[cell.row].height = desired_height
                adjusted += 1
    return adjusted


def process_detail_sheet(ctx: DetailSheetContext, order_context: OrderContext, stats: TransformStats) -> None:
    # 工分列处理：所有订单类型（含混油）均需填充工分公式
    if ctx.workpoint_col is None:
        stats.warnings.append(f"工作表 '{ctx.ws.title}' 缺少工分列，未补填工分")
    else:
        filled = fill_workpoint_formulas(
            ctx.ws,
            ctx.header_row,
            ctx.product_col,
            ctx.subtotal_col,
            ctx.workpoint_col,
        )
        stats.filled_workpoints += filled
        if filled > 0:
            stats.warnings.append(f"工作表 '{ctx.ws.title}' 补填了 {filled} 个工分公式")

    process_detail_rows(ctx, order_context, stats)
    cleanup_sheet_notes(ctx.ws, order_context, stats)
    stats.deleted_blank_rows += delete_blank_rows_between_data_and_summary(ctx.wb, ctx.ws, ctx.header_row)
    restore_color_block_merge(ctx.ws, ctx.header_row, ctx.product_col, ctx.color_col)
    fit_long_note_rows(ctx.ws)


EXTERNAL_REF_RE = re.compile(
    r"^=\+?'?\[(?P<book>\d+)\](?P<sheet>[^'!\[\]]+)'?!(?P<cell>\$?[A-Z]{1,3}\$?\d+)$"
)


def _cell_label(sheet_title: str, coord: str) -> str:
    """告警里指位子要说人话：数字页就说"第 3 页 L12"，别写 3!L12。"""
    return f"第 {sheet_title} 页 {coord}" if sheet_title.isdigit() else f"{sheet_title} 的 {coord}"


def external_link_cache(wb: Any) -> dict[tuple[int, str, str], Any]:
    """Collect cached cell values from the workbook's external links."""
    cache: dict[tuple[int, str, str], Any] = {}
    for book_idx, link in enumerate(wb._external_links, start=1):
        book = link.externalBook
        names = list(book.sheetNames.sheetName) if book.sheetNames else []
        dataset = book.sheetDataSet
        if not names or dataset is None:
            continue
        for sheet_data in dataset.sheetData:
            if sheet_data.sheetId is None or sheet_data.sheetId >= len(names):
                continue
            sheet_name = names[sheet_data.sheetId]
            for ext_row in sheet_data.row:
                for ext_cell in ext_row.cell:
                    value: Any = ext_cell.v
                    if value is not None and ext_cell.t != "str":
                        try:
                            value = int(value) if re.fullmatch(r"-?\d+", str(value)) else float(value)
                        except (TypeError, ValueError):
                            pass
                    cache[(book_idx, sheet_name, ext_cell.r)] = value
    return cache


def freeze_external_links(wb: Any, stats: TransformStats) -> int:
    """Replace external-workbook formulas with their cached values.

    Quote templates keep dead links to old template files on other machines
    (e.g. F:\\WeChat Files...). WPS/Excel then shows a "links cannot be
    updated" warning on every open. Freezing the cached values keeps the
    displayed content identical and lets us drop the external link parts.
    """
    if not wb._external_links:
        return 0
    cache = external_link_cache(wb)
    frozen = 0
    unresolved: list[str] = []
    for ws in wb.worksheets:
        max_col = min(ws.max_column, BUSINESS_SCAN_MAX_COL)
        for row in ws.iter_rows(max_col=max_col):
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("=") or "[" not in value:
                    continue
                match = EXTERNAL_REF_RE.match(value)
                if not match:
                    unresolved.append(_cell_label(ws.title, cell.coordinate))
                    continue
                key = (
                    int(match.group("book")),
                    match.group("sheet"),
                    match.group("cell").replace("$", ""),
                )
                cell.value = cache.get(key)
                frozen += 1
    if unresolved:
        shown = "、".join(unresolved[:10])
        more = f"（另外还有 {len(unresolved) - 10} 处）" if len(unresolved) > 10 else ""
        stats.warnings.append(
            f"有 {len(unresolved)} 个格子是从别的文件取数的，程序没法安全换算，只能保留原样：{shown}{more}。"
            "打开后这几格要是显示空白或 0，请照报价单手填。"
        )
    else:
        wb._external_links = []
    return frozen


def enable_excel_recalculation(wb: Any) -> None:
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass


def soffice_profile_uri(profile_dir: Path = SOFFICE_PROFILE_DIR) -> str:
    profile_dir.mkdir(parents=True, exist_ok=True)
    return profile_dir.resolve().as_uri()


def soffice_base_command(profile_dir: Path = SOFFICE_PROFILE_DIR) -> list[str]:
    return [
        "soffice",
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--norestore",
        "--nolockcheck",
        f"-env:UserInstallation={soffice_profile_uri(profile_dir)}",
    ]


def warm_xls_converter(timeout: int = 30) -> float:
    """Initialize LibreOffice's profile so the first real .xls conversion is not cold."""
    global _SOFFICE_WARMED
    if _SOFFICE_WARMED:
        return 0.0

    with _SOFFICE_LOCK:
        if _SOFFICE_WARMED:
            return 0.0

        start = time.perf_counter()
        result = subprocess.run(
            soffice_base_command() + ["--terminate_after_init"],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        elapsed = time.perf_counter() - start
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(f"LibreOffice profile warm-up failed: {detail}")
        _SOFFICE_WARMED = True
        return elapsed


FORMULA_CELL_RE = re.compile(
    r'<c r="(?P<ref>[A-Z]{1,3}[0-9]+)"(?P<attrs>[^>]*?)>'
    r'(?P<formula><f[^>]*?(?:/>|>.*?</f>))'
    r'(?P<cached><v[^>]*?(?:/>|>[^<]*</v>))?'
    r"</c>",
    re.S,
)
_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def collect_formula_cells(wb: Any) -> dict[str, list[str]]:
    """记下每个工作表里哪些单元格是公式，后面要按坐标回填缓存值。"""
    refs: dict[str, list[str]] = {}
    for ws in wb.worksheets:
        cells = [
            cell.coordinate
            for row in ws.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
        if cells:
            refs[ws.title] = cells
    return refs


def _format_number(value: int | float) -> str:
    if isinstance(value, int):
        return str(value)
    if float(value).is_integer() and abs(value) < 1e15:
        return str(int(value))
    return repr(float(value))


def _cached_xml_value(value: Any) -> tuple[str, str] | None:
    """把计算结果转成 (<v> 里的文本, 单元格 t 属性)；None 表示不写缓存。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return ("1" if value else "0", "b")
    if isinstance(value, (int, float)):
        return (_format_number(value), "")
    if isinstance(value, datetime.datetime):
        return (_format_number(to_excel_serial(value)), "")
    if isinstance(value, datetime.date):
        return (_format_number(to_excel_serial(datetime.datetime(value.year, value.month, value.day))), "")
    text = str(value)
    if text.startswith("#"):
        return (escape(text), "e")
    if not text:
        return None
    return (escape(text), "str")


def _sheet_part_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """工作表名 -> xl/worksheets/sheetN.xml。"""
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    parts: dict[str, str] = {}
    for sheet in workbook.iter(_MAIN_NS + "sheet"):
        target = targets.get(sheet.get(_REL_NS + "id"))
        if not target:
            continue
        parts[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else "xl/" + target
    return parts


def _read_computed_values(
    computed_path: Path, formula_refs: dict[str, list[str]]
) -> dict[str, dict[str, tuple[str, str]]]:
    wb = load_workbook(computed_path, data_only=True)
    values: dict[str, dict[str, tuple[str, str]]] = {}
    for name, refs in formula_refs.items():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        sheet_values: dict[str, tuple[str, str]] = {}
        for ref in refs:
            entry = _cached_xml_value(ws[ref].value)
            if entry is not None:
                sheet_values[ref] = entry
        if sheet_values:
            values[name] = sheet_values
    return values


def _inject_cached_values(
    workbook_path: Path, values: dict[str, dict[str, tuple[str, str]]]
) -> int:
    """把缓存值写进 xlsx 的 <v>，其余字节保持原样（图片、样式都不动）。"""
    if not values:
        return 0

    with zipfile.ZipFile(workbook_path) as zf:
        parts = _sheet_part_map(zf)
        entries = [(info, zf.read(info.filename)) for info in zf.infolist()]

    part_values = {
        parts[name]: sheet_values for name, sheet_values in values.items() if name in parts
    }
    filled = 0

    def replace(match: re.Match[str], sheet_values: dict[str, tuple[str, str]]) -> str:
        nonlocal filled
        entry = sheet_values.get(match.group("ref"))
        if entry is None:
            return match.group(0)
        text, cell_type = entry
        attrs = re.sub(r'\s+t="[^"]*"', "", match.group("attrs"))
        if cell_type:
            attrs = f'{attrs} t="{cell_type}"'
        filled += 1
        return f'<c r="{match.group("ref")}"{attrs}>{match.group("formula")}<v>{text}</v></c>'

    rewritten: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in entries:
        sheet_values = part_values.get(info.filename)
        if not sheet_values:
            rewritten.append((info, data))
            continue
        text = data.decode("utf-8")
        text = FORMULA_CELL_RE.sub(lambda m, sv=sheet_values: replace(m, sv), text)
        rewritten.append((info, text.encode("utf-8")))

    if not filled:
        return 0

    tmp_path = workbook_path.with_name(workbook_path.name + ".tmp")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in rewritten:
            out.writestr(info, data)
    tmp_path.replace(workbook_path)
    return filled


def embed_cached_values(
    output_path: Path, formula_refs: dict[str, list[str]]
) -> tuple[int, int]:
    """让 LibreOffice 重算一遍，把公式结果回填成 xlsx 的缓存值。

    openpyxl 保存时只写公式、不写结果，飞书网页版预览不会重算，公式列就整列空白。
    这里借容器里已有的 LibreOffice 算出结果，再按坐标写回原文件，避免整本重存
    导致图片/样式被改。失败时静默跳过（宁可没有缓存，也不能弄坏文件），
    由调用方根据返回的两个数字决定是否在群里提醒。

    返回 (算出来的格数, 实际写回文件的格数)。
    """
    if not CACHED_VALUES_ENABLED or not formula_refs or shutil.which("soffice") is None:
        return (0, 0)
    try:
        with tempfile.TemporaryDirectory(prefix="workshop_cached_") as tmp:
            cmd = soffice_base_command() + [
                "--convert-to",
                "xlsx",
                "--outdir",
                tmp,
                str(output_path),
            ]
            with _SOFFICE_LOCK:
                warm_xls_converter()
                result = subprocess.run(
                    cmd, capture_output=True, text=True, timeout=SOFFICE_TIMEOUT_SECONDS
                )
            if result.returncode != 0:
                return (0, 0)
            computed = Path(tmp) / f"{output_path.stem}.xlsx"
            if not computed.exists():
                candidates = list(Path(tmp).glob("*.xlsx"))
                if len(candidates) != 1:
                    return (0, 0)
                computed = candidates[0]
            values = _read_computed_values(computed, formula_refs)
        expected = sum(len(sheet_values) for sheet_values in values.values())
        return (expected, _inject_cached_values(output_path, values))
    except Exception:
        return (0, 0)


def transform(input_path: Path, output_path: Path, discount: float, order_type: str = "auto") -> dict[str, int | str]:
    temp_dir: tempfile.TemporaryDirectory[str] | None = None
    workbook_path = input_path
    if input_path.suffix.lower() == ".xls":
        temp_dir = tempfile.TemporaryDirectory(prefix="workshop_order_")
        workbook_path = Path(temp_dir.name) / "converted.xlsx"
        convert_xls_to_xlsx(input_path, workbook_path)

    wb = load_workbook(workbook_path)
    order_context = build_order_context(input_path, discount, order_type, wb)
    cached_wb = (
        load_workbook(workbook_path, data_only=True)
        if order_context.resolved_order_type == "tiepi" and workbook_needs_cached_prices(wb)
        else None
    )
    stats = TransformStats(order_type=order_context.resolved_order_type)

    for ws in wb.worksheets:
        detail_context = build_detail_sheet_context(wb, cached_wb, ws)
        if detail_context:
            process_detail_sheet(detail_context, order_context, stats)

    enable_excel_recalculation(wb)
    stats.moved_hardware_sheets = move_hardware_sheets_to_end(wb)
    stats.updated_order_numbers = normalize_order_numbers(wb, order_context.inferred_order_number)
    stats.frozen_external_links = freeze_external_links(wb, stats)

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()

    # 保存完再回填公式缓存值：飞书网页版/手机端预览不重算公式，
    # 没有 <v> 缓存就会整列显示空白。
    formula_refs = collect_formula_cells(wb)
    total_formulas = sum(len(refs) for refs in formula_refs.values())
    computed, stats.cached_values_embedded = embed_cached_values(output_path, formula_refs)

    # 算不出来的必须让群里知道：光看 Excel 是发现不了的（Excel 会自己重算），
    # 只有飞书网页版/手机预览才会露馅成空白。
    if total_formulas:
        if stats.cached_values_embedded == 0:
            stats.warnings.append(
                f"⚠️ 本单 {total_formulas} 个公式结果全部没算出来，飞书网页版/手机预览时这些格子是空白"
                "（Excel/WPS 打开正常）。请人工核对金额，或让拆单重新下单。"
            )
        elif stats.cached_values_embedded < computed:
            stats.warnings.append(
                f"⚠️ 公式结果只写回 {stats.cached_values_embedded}/{computed} 个，飞书网页版预览可能有个别空白格"
                "（Excel/WPS 打开正常），请翻一遍确认。"
            )
        elif computed * 2 < total_formulas:
            stats.warnings.append(
                f"⚠️ {total_formulas} 个公式只有 {computed} 个算出值，其余为正常空结果；"
                "若发现金额列空白请人工核对。"
            )
    return stats.to_dict()


def convert_xls_to_xlsx(input_path: Path, output_path: Path) -> None:
    """Convert legacy .xls to .xlsx while preserving workbook formatting.

    Docker/Linux cannot use Excel COM. LibreOffice headless keeps styles,
    merged cells, column widths, images and sheet layout far better than a
    pandas data-frame rewrite. If LibreOffice is unavailable or fails, raise a
    clear error instead of silently producing a data-only workbook.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="xls_convert_") as tmp:
        tmpdir = Path(tmp)
        cmd = soffice_base_command() + [
            "--convert-to",
            "xlsx",
            "--outdir",
            str(tmpdir),
            str(input_path),
        ]

        with _SOFFICE_LOCK:
            warm_xls_converter()
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=SOFFICE_TIMEOUT_SECONDS)
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(f"LibreOffice failed to convert .xls to .xlsx: {detail}")

        converted = tmpdir / f"{input_path.stem}.xlsx"
        if not converted.exists():
            candidates = list(tmpdir.glob("*.xlsx"))
            if len(candidates) == 1:
                converted = candidates[0]
            else:
                detail = (result.stderr or result.stdout or "").strip()
                raise RuntimeError(f"LibreOffice did not create an .xlsx file: {detail}")

        output_path.write_bytes(converted.read_bytes())


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}\u4e0b\u8f66\u95f4.xlsx")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert original order workbook to workshop order workbook.")
    parser.add_argument("input", type=Path, help="Original .xlsx workbook")
    parser.add_argument("output", type=Path, nargs="?", help="Output .xlsx workbook")
    parser.add_argument("--discount", type=float, default=DEFAULT_DISCOUNT, help="Unit-price multiplier, default: 0.85")
    parser.add_argument(
        "--order-type",
        choices=sorted(ORDER_TYPES),
        default="auto",
        help="Order rule set: auto detects hunyou only when a color-cell value starts with '混油'; tiepi uses full rules; hunyou only clears packaging and color fees.",
    )
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    output_path = (args.output or default_output_path(input_path)).expanduser().resolve()

    if input_path == output_path:
        raise SystemExit("Output path must be different from input path.")
    if not input_path.exists():
        raise SystemExit(f"Input workbook not found: {input_path}")

    stats = transform(input_path, output_path, args.discount, args.order_type)
    print(f"Saved: {output_path}")
    print(
        "Stats: "
        f"order_type={stats['order_type']}, "
        f"discounted_prices={stats['discounted_prices']}, "
        f"cleared_wood_boxes={stats['cleared_wood_boxes']}, "
        f"cleared_non_standard_color_fees={stats['cleared_non_standard_color_fees']}, "
        f"summary_constants_removed={stats['summary_constants_removed']}, "
        f"updated_page_notes={stats['updated_page_notes']}, "
        f"cleared_bottom_notes={stats['cleared_bottom_notes']}, "f"frozen_external_links={stats['frozen_external_links']}, "
        f"filled_workpoints={stats['filled_workpoints']}, "
        f"skipped_prices={stats['skipped_prices']}, "
        f"deleted_blank_rows={stats['deleted_blank_rows']}, "
        f"moved_hardware_sheets={stats['moved_hardware_sheets']}, "
        f"updated_order_numbers={stats['updated_order_numbers']}, "
        f"cached_values_embedded={stats['cached_values_embedded']}"
    )


if __name__ == "__main__":
    main()
