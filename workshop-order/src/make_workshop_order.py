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
import operator
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange


DEFAULT_DISCOUNT = 0.85

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
KW_TOTAL = "\u5408\u8ba1"
KW_BOARD = "\u677f\u6750"
KW_PAYMENT_TOTAL = "\u5408\u8ba1\u91d1\u989d"
KW_NEED_PAY = "\u9700\u652f\u4ed8"
KW_NEED_PAY_SHORT = "\u9700\u4ed8"
KW_DISCOUNTED = "\u6298\u540e"
KW_ORIGINAL_ORDER = "\u539f\u5355\u53f7"
KW_COLOR_DIFF = "\u8272\u5dee"
KW_AFTER_SALES = "\u552e\u540e"

DROP_NOTE_KEYWORDS = (
    KW_ORIGINAL_ORDER,
)

PAGE_TEXT_RE = re.compile(r"\u6b64\u5355\u5171\s*(\d+)\s*\u9875")
PAGE_COUNT_RE = re.compile(r"\u5171\u8ba1\s*(\d+)\s*\u9875")

# Payment-related keywords that should be stripped from bottom notes.
PAYMENT_AMOUNT_KEYWORDS = (
    "\u5408\u8ba1\u91d1\u989d",
    "85\u6298\u4f18\u60e0",
    "\u9700\u652f\u4ed8",
    "\u9700\u4ed8",
    "\u5e94\u652f\u4ed8",
    "\u5e94\u4ed8",
    "\u4f18\u60e0\u91d1\u989d",
    "\u6298\u540e\u91d1\u989d",
)


def remove_payment_amount_info(value: str) -> str:
    """Strip payment amount suffixes from bottom notes, keep after-sales/page text."""
    min_idx = len(value)
    for kw in PAYMENT_AMOUNT_KEYWORDS:
        idx = value.find(kw)
        if idx != -1 and idx < min_idx:
            min_idx = idx
    if min_idx < len(value):
        value = value[:min_idx].rstrip(",.，。；;:!：！ ")
    return value


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
    for row in ws.iter_rows():
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


def get_price_value(formula_cell: Any, cached_cell: Any) -> float | None:
    cached = cached_cell.value
    if is_number(cached):
        return float(cached)
    raw = formula_cell.value
    if is_number(raw):
        return float(raw)
    if isinstance(raw, str):
        return safe_eval_arithmetic(raw)
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


def detect_order_type(wb: Any, requested: str) -> str:
    if requested != "auto":
        return requested

    for ws in wb.worksheets:
        color_header = find_color_column(ws)
        if not color_header:
            continue
        header_row, color_col, product_col = color_header
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

            if color.startswith(KW_HUNYOU):
                return "hunyou"
    return "tiepi"


def find_color_column(ws: Any) -> tuple[int, int, int | None] | None:
    """Return header_row, color_col, product_col for production detail sheets."""
    for row in ws.iter_rows():
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
        for row in ws.iter_rows():
            for cell in row:
                value = text(cell.value)
                match = PAGE_TEXT_RE.search(value) or PAGE_COUNT_RE.search(value)
                if match:
                    return match.group(1)
    return None


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
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("="):
                    continue
                adjusted = adjust_formula_after_row_delete(value, ws.title, deleted_ws.title, deleted_row, amount)
                if adjusted != value:
                    cell.value = adjusted
                    changed += 1
    return changed


def prepare_merged_cells_for_row_delete(ws: Any, deleted_row: int, amount: int = 1) -> list[str]:
    shifted_ranges: list[str] = []
    for merged_range in list(ws.merged_cells.ranges):
        cell_range = CellRange(str(merged_range))
        if cell_range.min_row <= deleted_row <= cell_range.max_row:
            ws.unmerge_cells(str(merged_range))
        elif cell_range.min_row > deleted_row:
            ws.unmerge_cells(str(merged_range))
            cell_range.shift(row_shift=-amount)
            shifted_ranges.append(str(cell_range))
    return shifted_ranges


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

    for row_idx in range(data_start, ws.max_row + 1):
        row_labels = [text(ws.cell(row_idx, col).value) for col in range(1, product_col + 1)]
        if any(KW_WOOD_BOX in value or value == KW_TOTAL for value in row_labels):
            break

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
        if isinstance(workpoint_cell.value, str) and workpoint_cell.value.startswith("="):
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
        for col_idx in range(1, ws.max_column + 1):
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
        for col_idx in range(1, ws.max_column + 1):
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


def transform(input_path: Path, output_path: Path, discount: float, order_type: str = "auto") -> dict[str, int | str]:
    temp_dir: tempfile.TemporaryDirectory[str] | None = None
    workbook_path = input_path
    if input_path.suffix.lower() == ".xls":
        temp_dir = tempfile.TemporaryDirectory(prefix="workshop_order_")
        workbook_path = Path(temp_dir.name) / "converted.xlsx"
        convert_xls_to_xlsx(input_path, workbook_path)

    wb = load_workbook(workbook_path)
    cached_wb = load_workbook(workbook_path, data_only=True)
    resolved_order_type = detect_order_type(wb, order_type)
    page_count = extract_page_count(wb)

    stats = {
        "order_type": resolved_order_type,
        "discounted_prices": 0,
        "cleared_wood_boxes": 0,
        "cleared_non_standard_color_fees": 0,
        "updated_page_notes": 0,
        "summary_constants_removed": 0,
        "cleared_bottom_notes": 0,
        "filled_workpoints": 0,
        "skipped_prices": 0,
        "deleted_blank_rows": 0,
    }

    for ws in wb.worksheets:
        cached_ws = cached_wb[ws.title]
        headers = find_header_columns(ws)
        if headers:
            header_row, product_col, price_col, subtotal_col, workpoint_col = headers
            color_header = find_color_column(ws)
            color_col = color_header[1] if color_header else None
            if resolved_order_type == "tiepi":
                stats["filled_workpoints"] += fill_workpoint_formulas(
                    ws,
                    header_row,
                    product_col,
                    subtotal_col,
                    workpoint_col,
                )

            current_product = ""
            current_color = ""
            for row_idx in range(header_row + 1, ws.max_row + 1):
                raw_product = text(ws.cell(row_idx, product_col).value)
                price_cell = ws.cell(row_idx, price_col)
                cached_price_cell = cached_ws.cell(row_idx, price_col)
                subtotal_cell = ws.cell(row_idx, subtotal_col)

                if raw_product:
                    current_product = raw_product
                elif current_product and (price_cell.value not in (None, "") or subtotal_cell.value not in (None, "")):
                    # continuation of a vertically merged product cell
                    pass
                else:
                    current_product = ""

                if color_col:
                    raw_color = text(ws.cell(row_idx, color_col).value)
                    if raw_color:
                        current_color = raw_color
                    elif current_product and (price_cell.value not in (None, "") or subtotal_cell.value not in (None, "")):
                        # continuation of a vertically merged color cell
                        pass
                    else:
                        current_color = ""

                product = current_product
                color = current_color

                if is_color_fee_product(product):
                    last_clear_col = max(
                        subtotal_col,
                        workpoint_col or subtotal_col,
                        min(ws.max_column, 18),
                    )
                    changed = False
                    for col_idx in range(1, last_clear_col + 1):
                        row_cell = ws.cell(row_idx, col_idx)
                        if not isinstance(row_cell, MergedCell) and row_cell.value not in (None, ""):
                            row_cell.value = None
                            changed = True
                    if changed:
                        stats["cleared_non_standard_color_fees"] += 1
                    continue

                if KW_WOOD_BOX in product:
                    if subtotal_cell.value not in (None, ""):
                        subtotal_cell.value = None
                        stats["cleared_wood_boxes"] += 1
                    continue

                if not product or KW_TOTAL in product or isinstance(price_cell, MergedCell):
                    continue

                if resolved_order_type == "hunyou":
                    continue

                price = get_price_value(price_cell, cached_price_cell)
                if price is None:
                    if price_cell.value not in (None, ""):
                        stats["skipped_prices"] += 1
                    continue

                price_cell.value = price * discount
                stats["discounted_prices"] += 1

            for row in ws.iter_rows():
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
                            stats["summary_constants_removed"] += 1
                    elif any(keyword in value for keyword in DROP_NOTE_KEYWORDS):
                        cell.value = None
                        stats["cleared_bottom_notes"] += 1
                    elif any(kw in value for kw in PAYMENT_AMOUNT_KEYWORDS):
                        new_value = remove_payment_amount_info(value)
                        if new_value != value:
                            cell.value = new_value
                            stats["updated_page_notes"] += 1
                    elif resolved_order_type == "hunyou":
                        continue

            stats["deleted_blank_rows"] += delete_blank_rows_between_data_and_summary(wb, ws, header_row)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        return stats
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()


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
        cmd = [
            "soffice",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(tmpdir),
            str(input_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
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
        f"cleared_bottom_notes={stats['cleared_bottom_notes']}, "
        f"filled_workpoints={stats['filled_workpoints']}, "
        f"skipped_prices={stats['skipped_prices']}, "
        f"deleted_blank_rows={stats['deleted_blank_rows']}"
    )


if __name__ == "__main__":
    main()
