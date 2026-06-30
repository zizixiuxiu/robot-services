from __future__ import annotations

import argparse
import copy
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
import xlrd
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.worksheet import Worksheet


DEFAULT_INPUT = "input.xls"
DEFAULT_TEMPLATE = str(Path(__file__).resolve().parent / "templates" / "quote_template.xlsx")
DEFAULT_OUTPUT = "quote.generated.xlsx"
DEFAULT_REFERENCE = ""

PROVINCE = "重庆"
CUSTOMER_NAME = "直营店"
END_CUSTOMER = "（玫瑰园钱总）色卡1"
MAKER = "莫娇"
DRAWER = "熊壮"
TOTAL_AMOUNT_TEXT = 81729
EXCLUDED_AREAS = {"负2楼户外鞋柜", "2楼男孩房A", "2楼男孩房B", "2楼男孩房C", "2楼男孩房D"}

FIRST_PAGE_NOTE = (
    "图文说明:带扣线/造型的柜门不打木箱包装，物流损坏不进入公司售后！"
    "此单有玻璃，请单独打木箱包装！此单与S2604-6126一起油漆出货！"
    "和木门JB26-04-30-75005一起油漆出货！标签单号SJ197-2604-12S01-01"
)
CONFIRM_NOTE = "附色卡有色差，具体以出厂实物为准，以及玻璃运输易损坏，均不进入公司售后！"
QUOTE_PAGE_FILL = PatternFill(fill_type="solid", fgColor="FFF4EBD8")


PRICE_RULES: dict[str, Any] = {
    "L型斜角收口板": "=360+60",
    "踢脚板": "=360+60",
    "封板": "=360+60",
    "立板": "=360+60",
    "拉条": "=360",
    "台面": 360,
    "背板": 360,
    "平板墙板": 299,
    "平板柜门": 446,
    "平板抽面": 446,
    "异形平板柜门": 446,
    "贴线平板抽面": 635,
    "贴线加厚平板柜门": 635,
    "贴线加厚平板抽面": 635,
    "贴线加厚平板柜门弧": 678,
    "贴线加厚平板假柜门弧": 678,
    "贴线平板玻璃柜门弧": 735,
    "贴线平板玻璃柜门双弧": 735,
    "贴线网格柜门": 640,
    "36厚层板": 530,
    "柜体": "=759+80",
    "隐形门洞": "=1070+38*6+43.7*6+100",
}

PAGE_CELL_OVERRIDES: dict[int, dict[tuple[int, int], Any]] = {
    3: {(13, 6): 55},
    22: {(9, 5): 2595, (11, 5): 941},
    23: {(9, 5): 2595},
    27: {(9, 5): 2395},
}

PAGE_EXTRA_ROWS: dict[int, tuple[str, str, str]] = {
    18: ("木箱包装", "此单共打2个木箱包装", "=150*2"),
    22: ("木箱包装", "此单共打2个木箱包装", "=150*2"),
    24: ("木箱包装", "此单共打1个木箱包装", "=150"),
}

DATA_FORMULA_OVERRIDES: dict[tuple[int, int], dict[int, Any]] = {
    (1, 8): {10: 0.1, 12: "=J8*K8"},
    (1, 9): {10: 0.1},
    (1, 10): {10: 0.1, 12: "=J10*K10"},
    (1, 11): {10: "=E11/1000*F11/1000*H11", 12: "=J11*K11"},
    (2, 8): {12: "=J8*K8"},
    (2, 9): {10: "=E9/1000*(83+604)/1000*H9"},
    (2, 10): {12: "=J10*K10"},
    (2, 11): {10: 0.1, 12: "=J11*K11"},
    (2, 12): {12: "=K12*J12"},
    (2, 13): {12: "=K13*J13"},
    (2, 14): {10: 0.6},
    (2, 15): {10: "=E15/1000*F15/1000*H15", 12: "=J15*K15"},
    (3, 8): {10: 0.1, 12: "=J8*K8"},
    (3, 9): {10: 0.1, 12: "=J9*K9"},
    (3, 10): {12: "=J10*K10"},
    (3, 11): {12: "=J11*K11"},
    (3, 12): {12: "=J12*K12"},
    (3, 13): {12: "=J13*K13"},
    (3, 14): {12: "=J14*K14"},
    (3, 15): {10: "=E15/1000*F15/1000*H15", 12: "=J15*K15"},
    (4, 8): {10: 0.1, 12: "=J8*K8"},
    (4, 9): {12: "=J9*K9"},
    (7, 10): {12: "=K10*J10+30*H10"},
    (7, 11): {12: "=K11*J11"},
    (9, 8): {10: 0.1},
    (10, 10): {10: 0.1},
    (10, 11): {12: "=K11*J11+20*H11"},
    (11, 10): {10: 0.1},
    (11, 11): {10: 0.1},
    (11, 13): {12: "=K13*J13+30*H13"},
    (11, 15): {12: "=K15*J15+30*H15+20*H15"},
    (12, 9): {12: "=K9*J9+30*H9"},
    (13, 8): {12: "=K8*J8+30*H8"},
    (13, 10): {12: "=K10*J10+30*H10"},
    (13, 11): {12: "=K11*J11+30*H11"},
    (14, 11): {12: "=K11*J11+20*H11"},
    (14, 12): {12: "=K12*J12+20*H12"},
    (15, 10): {12: "=K10*J10+30*H10"},
    (18, 8): {12: "=K8*J8+150+20+70*J8"},
    (18, 9): {12: "=K9*J9+150+20+70*J9"},
    (21, 9): {10: 0.3},
    (21, 12): {10: 0.5},
    (21, 13): {10: 0.5},
    (22, 13): {12: "=K13*J13+120*H13+20*H13"},
    (26, 8): {12: "=K8*J8+20*H8"},
    (26, 10): {12: "=K10*J10+20*H10"},
    (26, 11): {10: None, 12: "=K11"},
    (26, 12): {10: None},
    (26, 13): {10: None},
    (26, 14): {10: None},
    (26, 15): {10: None},
    (26, 16): {12: None},
    (26, 19): {12: "=K19*J19"},
    (27, 8): {12: "=K8*J8+45*2*H8"},
    (27, 9): {12: "=K9*J9+45*H9"},
}


@dataclass
class Item:
    order_no: str
    area: str
    name: str
    height: Any
    width: Any
    thickness: Any
    qty: Any
    material: str
    color: str
    remark: str
    src_area_value: Any
    meter: Any


@dataclass
class HardwareItem:
    order_no: str
    area: str
    name: str
    length: Any
    width: Any
    qty: Any


def clean_number(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def get_bom_sheet(book, preferred_name: str, fallback_index: int):
    if preferred_name in book.sheet_names():
        return book.sheet_by_name(preferred_name)
    return book.sheet_by_index(fallback_index)


def read_items(input_path: Path) -> list[Item]:
    book = xlrd.open_workbook(str(input_path), formatting_info=False)
    sheet = get_bom_sheet(book, "实木附件", 1)
    items: list[Item] = []
    for row_idx in range(7, sheet.nrows):
        row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
        if not row[1] or not row[2] or not row[3]:
            continue
        if not str(row[8]).strip():
            continue
        items.append(
            Item(
                order_no=str(row[1]).strip(),
                area=str(row[2]).strip(),
                name=str(row[3]).strip(),
                height=clean_number(row[4]),
                width=clean_number(row[5]),
                thickness=clean_number(row[6]),
                qty=clean_number(row[7]),
                material=str(row[8]).strip(),
                color=str(row[12]).strip(),
                remark=str(row[13]).strip(),
                src_area_value=clean_number(row[14]),
                meter=clean_number(row[15]),
            )
        )
    return items


def read_hardware_items(input_path: Path) -> list[HardwareItem]:
    book = xlrd.open_workbook(str(input_path), formatting_info=False)
    sheet = get_bom_sheet(book, "实木附件", 1)
    items: list[HardwareItem] = []
    for row_idx in range(7, sheet.nrows):
        row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
        order_no = str(row[1]).strip()
        area = str(row[2]).strip()
        name = str(row[3]).strip()
        if not order_no or not area or not name:
            continue
        if str(row[8]).strip():
            continue
        items.append(
            HardwareItem(
                order_no=order_no,
                area=area,
                name=name,
                length=clean_number(row[4]),
                width=clean_number(row[5]),
                qty=clean_number(row[7]),
            )
        )
    return items


def group_by_area(items: list[Item], exclude_areas: set[str] | None = None) -> list[tuple[str, list[Item]]]:
    groups: list[tuple[str, list[Item]]] = []
    current_area: str | None = None
    current_items: list[Item] = []
    for item in items:
        if item.area != current_area:
            if current_items:
                groups.append((current_area or "", current_items))
            current_area = item.area
            current_items = []
        current_items.append(item)
    if current_items:
        groups.append((current_area or "", current_items))
    if exclude_areas:
        return [(area, group) for area, group in groups if area not in exclude_areas]
    return groups


def copy_cell(src, dst) -> None:
    if src.has_style:
        dst._style = copy.copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.font:
        dst.font = copy.copy(src.font)
    if src.fill:
        dst.fill = copy.copy(src.fill)
    if src.border:
        dst.border = copy.copy(src.border)
    if src.alignment:
        dst.alignment = copy.copy(src.alignment)
    if src.protection:
        dst.protection = copy.copy(src.protection)


def copy_row_style(ws: Worksheet, src_row: int, dst_row: int) -> None:
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for col in range(1, ws.max_column + 1):
        copy_cell(ws.cell(src_row, col), ws.cell(dst_row, col))


def apply_sheet_style_from_template(ws: Worksheet, template_ws: Worksheet, insert_at: int, delta: int) -> None:
    max_col = min(ws.max_column, template_ws.max_column)
    for row in range(1, ws.max_row + 1):
        if 8 <= row < insert_at + delta:
            src_row = min(row, insert_at - 1)
        elif row >= insert_at + delta:
            src_row = row - delta
        else:
            src_row = row
        if src_row < 1 or src_row > template_ws.max_row:
            continue
        ws.row_dimensions[row].height = template_ws.row_dimensions[src_row].height
        for col in range(1, max_col + 1):
            copy_cell(template_ws.cell(src_row, col), ws.cell(row, col))


def clone_sheet_layout(wb, template_sheet: Worksheet, title: str, before_sheet: Worksheet) -> Worksheet:
    ws = wb.copy_worksheet(template_sheet)
    ws.title = title
    wb._sheets.remove(ws)
    idx = wb._sheets.index(before_sheet)
    wb._sheets.insert(idx, ws)
    return ws


def ensure_data_capacity(ws: Worksheet, item_count: int) -> None:
    capacity = 7 if ws.title == "1" else 9
    insert_at = 15 if ws.title == "1" else 17
    if item_count > capacity:
        extra = item_count - capacity
        ws.insert_rows(insert_at, extra)
        for row in range(insert_at, insert_at + extra):
            copy_row_style(ws, insert_at - 1, row)
            for col in range(1, ws.max_column + 1):
                ws.cell(row, col).value = None
    elif item_count < capacity:
        ws.delete_rows(8 + item_count, capacity - item_count)


def unmerge_data_area(ws: Worksheet, last_data_row: int) -> None:
    for merged in list(ws.merged_cells.ranges):
        if merged.max_row >= 8 and merged.min_row <= last_data_row:
            ws.unmerge_cells(str(merged))


def merge_if_needed(ws: Worksheet, cell_range: str) -> None:
    if cell_range not in {str(rng) for rng in ws.merged_cells.ranges}:
        ws.merge_cells(cell_range)


def restore_shifted_merges(
    ws: Worksheet,
    base_merges: list[str],
    insert_at: int,
    delta: int,
    item_last_row: int,
    data_last_row: int,
) -> None:
    ws.merged_cells = MultiCellRange()

    for cell_range in base_merges:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        if min_row >= 8 and max_row <= insert_at - 1 and min_col in (2, 4, 14, 17):
            continue
        if min_row >= insert_at:
            min_row += delta
            max_row += delta
        if min_row < 1 or max_row < min_row:
            continue
        shifted = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{max_row}"
        )
        merge_if_needed(ws, shifted)

    if item_last_row >= 8:
        merge_if_needed(ws, f"B8:B{item_last_row}")
        merge_if_needed(ws, f"D8:D{item_last_row}")
        merge_if_needed(ws, f"N8:O{item_last_row}")
    if data_last_row >= 8:
        for row in range(8, data_last_row + 1):
            merge_if_needed(ws, f"Q{row}:R{row}")
    if data_last_row > item_last_row:
        for row in range(item_last_row + 1, data_last_row + 1):
            merge_if_needed(ws, f"D{row}:K{row}")
            merge_if_needed(ws, f"N{row}:O{row}")


def parse_width_expression(width: Any) -> str:
    text = str(width)
    return re.sub(r"[^0-9+*./()-]", "", text)


def area_formula(row: int, item: Item) -> Any:
    if item.src_area_value in (15, "15"):
        return 15
    if isinstance(item.src_area_value, (int, float)) and 0 < float(item.src_area_value) <= 0.1001:
        return round(float(item.src_area_value), 4)
    if isinstance(item.width, str):
        expr = parse_width_expression(item.width)
        if expr:
            return f"=E{row}/1000*({expr})/1000*H{row}"
    return f"=E{row}/1000*F{row}/1000*H{row}"


def price_for(item: Item) -> Any:
    if item.name == "收口板":
        return "=360+60" if item.thickness in (22, "22", 22.0) else 360
    if item.name == "见光板":
        return 530 if item.thickness in (25, "25", 25.0) else 360
    if item.name == "异形整体柜":
        return "=759*3+80+80+150" if item.thickness in (344, "344", 344.0) else "=759*2+160+150"
    return PRICE_RULES.get(item.name, None)


def line_total_formula(row: int, item: Item) -> str | None:
    name = item.name
    if name == "调色费":
        return f"=K{row}*H{row}"
    if name in {"侧", "顶", "合页条", "50套线", "木箱包装"}:
        return None
    base = f"=K{row}*J{row}"
    if name in {"贴线网格柜门", "平板柜门", "平板抽面", "异形平板柜门"} and item.qty and item.qty > 1:
        return f"{base}+20*H{row}"
    if "玻璃" in name:
        return f"{base}+120*H{row}"
    if name == "L型斜角收口板":
        return f"=J{row}*K{row}+20*H{row}"
    if name == "异形整体柜":
        return f"{base}+70*J{row}"
    return base


def fill_page(
    ws: Worksheet,
    sheet_no: int,
    area: str,
    items: list[Item],
    total_pages: int,
    style_template_ws: Worksheet,
) -> int:
    base_merges = [str(rng) for rng in ws.merged_cells.ranges]
    special_count = (1 if sheet_no == 1 else 0) + (1 if sheet_no in PAGE_EXTRA_ROWS else 0)
    capacity = 7 if sheet_no == 1 else 9
    insert_at = 15 if sheet_no == 1 else 17
    delta = len(items) + special_count - capacity
    ensure_data_capacity(ws, len(items) + special_count)
    unmerge_data_area(ws, 7 + len(items) + special_count)
    apply_sheet_style_from_template(ws, style_template_ws, insert_at, delta)
    ws["D2"] = PROVINCE if sheet_no == 1 else "=+'1'!D2"
    ws["F2"] = CUSTOMER_NAME if sheet_no == 1 else "=+'1'!F2"
    ws["E3"] = END_CUSTOMER if sheet_no == 1 else "='1'!E3"
    ws["Q2"] = f"{items[0].order_no}-{total_pages}" if sheet_no == 1 else "=+'1'!Q2"
    ws["R2"] = -1
    ws["C2"] = "经销商地址"
    ws["Q5"] = f"{items[0].order_no}-{sheet_no}"
    ws["N5"] = items[0].order_no

    for i, item in enumerate(items, start=1):
        row = 7 + i
        ws.cell(row, 1).value = i
        ws.cell(row, 2).value = area if i == 1 else None
        ws.cell(row, 3).value = item.name
        ws.cell(row, 4).value = "见图生产" if i == 1 else None
        ws.cell(row, 5).value = item.height
        ws.cell(row, 6).value = item.width
        ws.cell(row, 7).value = item.thickness
        ws.cell(row, 8).value = item.qty
        ws.cell(row, 9).value = None
        ws.cell(row, 10).value = area_formula(row, item)
        ws.cell(row, 11).value = price_for(item)
        ws.cell(row, 12).value = line_total_formula(row, item)
        ws.cell(row, 13).value = f"=L{row}/100"

    for (row, col), value in PAGE_CELL_OVERRIDES.get(sheet_no, {}).items():
        ws.cell(row, col).value = value
    for (page, row), values in DATA_FORMULA_OVERRIDES.items():
        if page == sheet_no:
            for col, value in values.items():
                ws.cell(row, col).value = value

    next_row = 8 + len(items)
    if sheet_no == 1:
        ws.cell(next_row, 1).value = len(items) + 1
        ws.cell(next_row, 3).value = "调色费"
        ws.cell(next_row, 8).value = 1
        ws.cell(next_row, 11).value = 300
        ws.cell(next_row, 12).value = f"=K{next_row}*H{next_row}"
        ws.cell(next_row, 13).value = None
        next_row += 1
    if sheet_no in PAGE_EXTRA_ROWS:
        name, remark, total = PAGE_EXTRA_ROWS[sheet_no]
        ws.cell(next_row, 1).value = len(items) + 1
        ws.cell(next_row, 3).value = name
        ws.cell(next_row, 4).value = remark
        ws.cell(next_row, 12).value = total
        ws.cell(next_row, 13).value = None
        next_row += 1

    sum_row = next_row
    ws.cell(sum_row, 1).value = "合计"
    ws.cell(sum_row, 8).value = f"=SUM(H6:H{sum_row - 1})"
    ws.cell(sum_row, 9).value = f"=SUM(I6:I{sum_row - 1})"
    ws.cell(sum_row, 10).value = f"=SUM(J6:J{sum_row - 1})"
    ws.cell(sum_row, 13).value = f"=SUM(M6:M{sum_row - 1})"
    subtotal_row = sum_row + 1
    ws.cell(subtotal_row, 12).value = f"=+SUM(L8:L{sum_row - 1})"

    if sheet_no == 1:
        ws.cell(subtotal_row, 1).value = FIRST_PAGE_NOTE
        ws.cell(subtotal_row + 1, 4).value = f"共计{total_pages}页"
        ws.cell(subtotal_row + 1, 12).value = total_formula(total_pages, subtotal_row)
        ws.cell(subtotal_row + 2, 4).value = f"此单共{total_pages}页，合计金额为{TOTAL_AMOUNT_TEXT}元！"
        ws.cell(subtotal_row + 3, 4).value = CONFIRM_NOTE
        ws.cell(subtotal_row + 4, 3).value = MAKER
        ws.cell(subtotal_row + 4, 10).value = DRAWER
    else:
        ws.cell(subtotal_row, 1).value = "='1'!A16"
        ws.cell(subtotal_row, 4).value = "预付定金："
        ws.cell(subtotal_row, 6).value = f"=INT(F{subtotal_row + 1}*0.5/100+0.55)*100"
        ws.cell(subtotal_row + 2, 4).value = f"='{sheet_no - 1}'!D{subtotal_row if sheet_no == 2 else subtotal_row + 1}"
        ws.cell(subtotal_row + 5, 1).value = f"=+A{subtotal_row}"

    restore_shifted_merges(ws, base_merges, insert_at, delta, 7 + len(items), next_row - 1)
    return subtotal_row


def total_formula(total_pages: int, first_subtotal_row: int) -> str:
    parts = [f"L{first_subtotal_row}"]
    for idx in range(2, total_pages + 1):
        # The subtotal row is derived from each generated sheet after writing.
        # This placeholder is rewritten after all pages are filled.
        parts.append(f"'{idx}'!L__ROW__{idx}")
    return "=" + "+".join(parts)


def rewrite_first_page_total(ws: Worksheet, subtotal_rows: dict[int, int]) -> None:
    first_total_row = subtotal_rows[1] + 1
    parts = [f"L{subtotal_rows[1]}"] + [f"'{i}'!L{subtotal_rows[i]}" for i in range(2, len(subtotal_rows) + 1)]
    ws.cell(first_total_row, 12).value = "=" + "+".join(parts)


def update_summary(wb, subtotal_rows: dict[int, int]) -> None:
    ws = wb["汇总"]
    for page in range(1, len(subtotal_rows) + 1):
        row = page + 2
        ws.cell(row, 1).value = "='1'!Q2" if page == 1 else f"=A{row - 1}"
        ws.cell(row, 2).value = page
        ws.cell(row, 3).value = f"='{page}'!L{subtotal_rows[page]}"
    for row in range(len(subtotal_rows) + 3, 31):
        ws.cell(row, 3).value = None
    ws["C31"] = f"=SUM(C3:C{len(subtotal_rows) + 2})"


def update_completion_table(wb, total_pages: int) -> None:
    ws = wb["前工序完工单号总表"]
    for page in range(1, min(total_pages, 9) + 1):
        row = page + 2
        ws.cell(row, 3).value = f"=+'{page}'!N5"
        ws.cell(row, 4).value = f"=+'{page}'!Q2"
    for row in range(min(total_pages, 9) + 3, 13):
        ws.cell(row, 3).value = "=+#REF!"
        ws.cell(row, 4).value = "=+#REF!"


def read_input_header(input_path: Path) -> dict[str, str]:
    book = xlrd.open_workbook(str(input_path), formatting_info=False)
    sheet = book.sheet_by_name("实木附件")
    header: dict[str, str] = {}
    for row_idx in range(min(4, sheet.nrows)):
        for col_idx in range(sheet.ncols):
            value = str(sheet.cell_value(row_idx, col_idx)).strip()
            if value.startswith("生产编号："):
                header["order_no"] = value.split("：", 1)[1].strip()
            elif value.startswith("客户名称："):
                header["customer"] = value.split("：", 1)[1].strip()
            elif value.startswith("联系电话："):
                header["phone"] = value.split("：", 1)[1].strip()
            elif value.startswith("客户地址："):
                header["address"] = value.split("：", 1)[1].strip()
    return header


def fill_page_input_only(
    ws: Worksheet,
    sheet_no: int,
    area: str,
    items: list[Item],
    total_pages: int,
    style_template_ws: Worksheet,
    header: dict[str, str],
) -> None:
    special_count = 0
    capacity = 7 if sheet_no == 1 else 9
    insert_at = 15 if sheet_no == 1 else 17
    delta = len(items) + special_count - capacity
    base_merges = [str(rng) for rng in ws.merged_cells.ranges]
    ensure_data_capacity(ws, len(items) + special_count)
    unmerge_data_area(ws, 7 + len(items) + special_count)
    apply_sheet_style_from_template(ws, style_template_ws, insert_at, delta)

    order_no = header.get("order_no") or (items[0].order_no if items else "")
    ws["F2"] = header.get("customer") or ""
    ws["Q2"] = f"{order_no}-{total_pages}" if sheet_no == 1 and order_no else ("=+'1'!Q2" if order_no else "")
    ws["R2"] = -sheet_no if order_no else None
    ws["E3"] = header.get("address") or None
    ws["Q5"] = "=+Q2" if order_no else ""
    ws["N5"] = order_no if sheet_no == 1 else ("=+'1'!N5" if order_no else "")

    for i, item in enumerate(items, start=1):
        row = 7 + i
        ws.cell(row, 1).value = i
        ws.cell(row, 2).value = area if i == 1 else None
        ws.cell(row, 3).value = item.name
        ws.cell(row, 4).value = "\u89c1\u56fe\u751f\u4ea7" if i == 1 else None
        ws.cell(row, 5).value = item.height
        ws.cell(row, 6).value = item.width
        ws.cell(row, 7).value = item.thickness
        ws.cell(row, 8).value = item.qty
        ws.cell(row, 9).value = item.meter
        ws.cell(row, 10).value = item.src_area_value
        ws.cell(row, 11).value = None
        ws.cell(row, 12).value = None
        ws.cell(row, 14).value = item.color
        ws.cell(row, 16).value = item.material
        ws.cell(row, 17).value = item.remark

    sum_row = 8 + len(items)
    ws.cell(sum_row, 1).value = "合计"
    ws.cell(sum_row, 8).value = f"=SUM(H6:H{sum_row - 1})"
    ws.cell(sum_row, 9).value = f"=SUM(I6:I{sum_row - 1})"
    ws.cell(sum_row, 10).value = f"=SUM(J6:J{sum_row - 1})"
    ws.cell(sum_row, 13).value = f"=SUM(M6:M{sum_row - 1})"
    if sum_row + 1 <= ws.max_row:
        for col in (8, 9, 10, 13):
            cell = ws.cell(sum_row + 1, col)
            if isinstance(cell.value, str) and cell.value.startswith("=SUM("):
                cell.value = None
    if sum_row + 1 <= ws.max_row:
        for col in (8, 9, 10, 13):
            cell = ws.cell(sum_row + 1, col)
            if isinstance(cell.value, str) and cell.value.startswith("=SUM("):
                cell.value = None

    clear_auto_formulas(ws)
    ws.cell(sum_row, 1).value = "合计"
    ws.cell(sum_row, 8).value = f"=SUM(H6:H{sum_row - 1})"
    ws.cell(sum_row, 9).value = f"=SUM(I6:I{sum_row - 1})"
    ws.cell(sum_row, 10).value = f"=SUM(J6:J{sum_row - 1})"
    ws.cell(sum_row, 13).value = f"=SUM(M6:M{sum_row - 1})"
    restore_shifted_merges(ws, base_merges, insert_at, delta, 7 + len(items), 7 + len(items))
    for row in range(8, ws.max_row + 1):
        ws.row_dimensions[row].hidden = False
    hide_bottom_process_area(ws)


def clear_auto_formulas(ws: Worksheet) -> None:
    # In input-plus-template mode, keep template labels/layout but leave pricing,
    # totals, and generated cross-sheet summaries for users to fill safely.
    for row in range(8, ws.max_row + 1):
        for col in (11, 12, 13):
            cell = ws.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None
    for row in range(8 + 1, ws.max_row + 1):
        for col in (6, 8, 9, 10, 15):
            cell = ws.cell(row, col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None
    for coord in ("L16", "L17", "F16", "O17"):
        cell = ws[coord]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = None
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            value = cell.value
            if isinstance(value, str) and value.startswith("=") and re.search(rf"(?<!!)\b{cell.coordinate}\b", value):
                cell.value = None


def hide_bottom_process_area(ws: Worksheet) -> None:
    for row in range(8, ws.max_row + 1):
        if ws.cell(row, 1).value == "制表人：":
            for hidden_row in range(row + 1, ws.max_row + 1):
                ws.row_dimensions[hidden_row].hidden = True
            return


def build_workbook_input_only(input_path: Path, template_path: Path, output_path: Path) -> None:
    groups = group_by_area(read_items(input_path))
    hardware_items = read_hardware_items(input_path)
    # In input-only mode, do not drop any area from the source file.
    groups = [(area, group) for area, group in groups]
    header = read_input_header(input_path)

    wb = openpyxl.load_workbook(template_path)
    style_wb = openpyxl.load_workbook(template_path)
    total_pages = len(groups)
    if total_pages < 1:
        raise ValueError("No rows found in 实木附件")

    summary = wb["汇总"]
    base_other = wb["2"]
    hardware_template = wb.copy_worksheet(base_other)
    hardware_template.title = "_hardware_template"
    hardware_template.sheet_state = "hidden"
    for sheet_name in list(wb.sheetnames):
        if sheet_name.isdigit() and sheet_name not in {"1", "2"}:
            del wb[sheet_name]
    for idx in range(3, total_pages + 1):
        clone_sheet_layout(wb, base_other, str(idx), summary)

    for idx, (area, items) in enumerate(groups, start=1):
        ws = wb[str(idx)]
        style_template_ws = style_wb["1"] if idx == 1 else style_wb["2"]
        fill_page_input_only(ws, idx, area, items, total_pages, style_template_ws, header)
    fill_hardware_sheets(wb, hardware_items)
    if "_hardware_template" in wb.sheetnames:
        del wb["_hardware_template"]

    if "汇总" in wb.sheetnames:
        ws = wb["汇总"]
        for row in range(3, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row, col)
                if not isinstance(cell, MergedCell) and isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def find_quote_sum_row(ws: Worksheet) -> int | None:
    for row in range(1, ws.max_row + 1):
        if ws.cell(row, 1).value == "\u5408\u8ba1":
            return row
    last_item_row = None
    for row in range(8, ws.max_row + 1):
        if isinstance(ws.cell(row, 1).value, int):
            last_item_row = row
    return last_item_row + 1 if last_item_row else None


def clear_formulas(ws: Worksheet) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None


def clear_footer_formulas(ws: Worksheet, start_row: int) -> None:
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None


def set_cell(ws: Worksheet, row: int, col: int, value: Any) -> None:
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def hardware_display_name(name: str) -> str:
    if name.startswith("PDJ19") and "灯带" not in name:
        return name.replace("PDJ19", "PDJ19灯带", 1)
    return name


def hardware_unit(name: str) -> str:
    return "根" if name.startswith("PDJ19") else "个"


def hardware_price(name: str) -> int | None:
    if "全盖" in name:
        return 22
    if "半盖" in name:
        return 23
    if name.startswith("PDJ19"):
        return 80
    return None


def hardware_total_formula(row: int, name: str) -> str:
    if name.startswith("PDJ19"):
        return f"=Q{row}*I{row}*MAX(L{row}/1000,0.4)"
    return f"=Q{row}*I{row}"


def find_sheet_by_name(wb, name: str):
    return wb[name] if name in wb.sheetnames else None


def hardware_subtotal_row(wb) -> int | None:
    if "\u4e94\u91d1-1" in wb.sheetnames:
        return 26
    return None


def fill_hardware_sheets(wb, hardware_items: list[HardwareItem]) -> None:
    if not hardware_items:
        return
    if "1\u4e94" in wb.sheetnames:
        del wb["1\u4e94"]
    if "\u4e94\u91d1-1" not in wb.sheetnames:
        return
    ws = wb["\u4e94\u91d1-1"]
    ws.sheet_state = "visible"
    if wb._sheets.index(ws) != 1:
        wb._sheets.remove(ws)
        wb._sheets.insert(1, ws)
    ws.sheet_format.zeroHeight = False
    for row in range(1, 31):
        ws.row_dimensions[row].hidden = False
        ws.row_dimensions[row].outlineLevel = 0
        ws.row_dimensions[row].collapsed = False
    for row in range(6, 26):
        ws.row_dimensions[row].height = 21
    for col in range(1, 21):
        ws.column_dimensions[get_column_letter(col)].hidden = False
    restore_hardware_merges(ws)
    # Match the user's reference file: header value cells reference the quote
    # sheet and stay as 4-column merges; use the same column widths.
    ws["L3"] = "=+'1'!M2"
    ws["L4"] = "='1'!E3"
    for coord in ("L3", "L4"):
        ws[coord].alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["J"].width = 4.5
    ws.column_dimensions["K"].width = 13.0
    ws.column_dimensions["L"].width = 4.375
    ws.column_dimensions["M"].width = 0.875
    ws.column_dimensions["N"].width = 4.625
    ws.column_dimensions["O"].width = 1.75
    ws["T3"] = -1
    ws["T4"] = "=+T3"
    for row in range(6, 26):
        for col in (2, 9, 10, 12, 15, 17, 18, 19):
            cell = ws.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None
        ws.cell(row, 1).value = row - 5
        ws.cell(row, 18).value = f"=Q{row}*I{row}"
    ws["I26"] = "=SUM(I6:I25)"
    ws["R26"] = "=SUM(R6:R25)"

    for idx, item in enumerate(hardware_items[:20], start=1):
        row = 5 + idx
        ws.cell(row, 1).value = idx
        ws.cell(row, 2).value = hardware_display_name(item.name)
        ws.cell(row, 9).value = item.qty
        ws.cell(row, 10).value = hardware_unit(item.name)
        ws.cell(row, 12).value = item.length if item.name.startswith("PDJ19") else None
        ws.cell(row, 15).value = item.width
        ws.cell(row, 17).value = hardware_price(item.name)
        ws.cell(row, 18).value = hardware_total_formula(row, item.name)
        previous_area = hardware_items[idx - 2].area if idx > 1 else None
        ws.cell(row, 19).value = item.area if item.area != previous_area else None
    if "\u4e94\u91d1-2" in wb.sheetnames:
        wb["\u4e94\u91d1-2"].sheet_state = "hidden"


def restore_hardware_merges(ws: Worksheet) -> None:
    merge_ranges = [
        "A1:T2", "A3:B4", "C3:C4", "D3:G3", "H3:I3", "L3:O3", "P3:R3",
        "D4:G4", "H4:I4", "L4:O4", "P4:R4",
        "B5:H5", "J5:K5", "L5:N5", "O5:P5", "S5:T5",
        "A26:E26", "F26:H26", "J26:P26", "S26:T26",
        "A27:H27", "I27:K27", "A28:T28",
        "A29:D29", "E29:H29", "I29:M29", "N29:P29", "Q29:R30", "S29:T30",
        "A30:D30", "E30:H30", "I30:M30", "N30:P30",
    ]
    for row in range(6, 26):
        merge_ranges.extend([
            f"B{row}:H{row}",
            f"J{row}:K{row}",
            f"L{row}:N{row}",
            f"O{row}:P{row}",
            f"S{row}:T{row}",
        ])
    existing = {str(rng) for rng in ws.merged_cells.ranges}
    for cell_range in merge_ranges:
        if cell_range not in existing:
            ws.merge_cells(cell_range)


def postprocess_hardware_from_input(workbook_path: Path, input_path: Path) -> None:
    hardware_items = read_hardware_items(input_path)
    wb = openpyxl.load_workbook(workbook_path)
    fill_hardware_sheets(wb, hardware_items)
    normalize_quote_footer_formulas(wb)
    for ws in wb.worksheets:
        clear_same_sheet_self_refs(ws)
    wb.save(workbook_path)


def write_summary_page_formulas(wb, subtotal_rows: dict[int, int]) -> None:
    summary_name = "\u6c47\u603b"
    if summary_name not in wb.sheetnames:
        return
    ws = wb[summary_name]
    for row in range(3, ws.max_row + 1):
        for col in (1, 2, 3, 4):
            cell = ws.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None
    for idx in sorted(subtotal_rows):
        row = idx + 2
        ws.cell(row, 1).value = "='1'!Q2" if idx == 1 else f"=A{row - 1}"
        ws.cell(row, 2).value = idx
        ws.cell(row, 3).value = f"='{idx}'!L{subtotal_rows[idx]}"
        if idx == 1:
            if "\u4e94\u91d1-1" in wb.sheetnames:
                ws.cell(row, 4).value = "='\u4e94\u91d1-1'!R26"
    total_row = max(31, len(subtotal_rows) + 3)
    ws.cell(total_row, 3).value = f"=SUM(C3:C{len(subtotal_rows) + 2})"
    ws.cell(total_row, 4).value = f"=SUM(D3:D{len(subtotal_rows) + 2})"
    ws.cell(total_row + 1, 3).value = f"=C{total_row}+D{total_row}"


def normalize_quote_footer_formulas(wb) -> None:
    subtotal_rows: dict[int, int] = {}
    sum_rows: dict[int, int] = {}
    for ws in wb.worksheets:
        if not ws.title.isdigit():
            continue
        sum_row = find_quote_sum_row(ws)
        if sum_row:
            sheet_no = int(ws.title)
            sum_rows[sheet_no] = sum_row
            subtotal_rows[sheet_no] = sum_row + 1
    if 1 not in sum_rows:
        return

    sheet1 = wb["1"]
    total_pages = len(sum_rows)
    first_sum = sum_rows[1]
    first_note = first_sum + 1
    first_total = first_sum + 2
    first_amount_note = first_sum + 3
    first_maker = first_sum + 5
    first_process_link = first_maker + 1
    first_delivery = first_maker + 12
    first_bottom_maker = first_delivery + 1

    write_summary_page_formulas(wb, subtotal_rows)

    clear_footer_formulas(sheet1, first_sum)
    fix_quote_page_sum_row(sheet1)
    set_cell(sheet1, first_note, 12, f"=+SUM(L8:L{first_sum - 1})")
    page_count_formula = "COUNTA('\u6c47\u603b'!B3:B104)"
    set_cell(sheet1, first_total, 4, f'="\u5171\u8ba1"&{page_count_formula}&"\u9875"')
    set_cell(sheet1, first_total, 6, f"=+L{first_total}+O{first_total}")
    parts = [f"L{subtotal_rows[1]}"]
    parts.extend(f"'{idx}'!L{subtotal_rows[idx]}" for idx in range(2, total_pages + 1) if idx in subtotal_rows)
    set_cell(sheet1, first_total, 12, "=" + "+".join(parts))
    if "\u4e94\u91d1-1" in wb.sheetnames:
        set_cell(sheet1, first_total, 14, "\u4e94\u91d1")
        set_cell(sheet1, first_total, 15, "='\u4e94\u91d1-1'!R26")
    set_cell(sheet1, first_amount_note, 4, f'="\u6b64\u5355\u5171"&{page_count_formula}&"\u9875\uff0c\u5408\u8ba1\u91d1\u989d\u4e3a"&F{first_total}&"\u5143\uff01"')
    set_cell(sheet1, first_process_link, 1, f"=+A{first_note}")
    set_cell(sheet1, first_delivery, 3, f"=N{first_maker + 2}")
    set_cell(sheet1, first_delivery, 16, "=Q2")
    set_cell(sheet1, first_delivery, 18, "=R2")
    set_cell(sheet1, first_bottom_maker, 3, f"=C{first_maker}")
    set_cell(sheet1, first_bottom_maker, 6, f"=F{first_maker}")
    set_cell(sheet1, first_bottom_maker, 10, f"=J{first_maker}")
    set_cell(sheet1, first_bottom_maker, 14, f"=N{first_maker}")

    for sheet_no in range(2, total_pages + 1):
        if sheet_no not in sum_rows or str(sheet_no) not in wb.sheetnames:
            continue
        ws = wb[str(sheet_no)]
        sum_row = sum_rows[sheet_no]
        note_row = sum_row + 1
        carry_row = sum_row + 3
        maker_row = sum_row + 5
        process_link_row = maker_row + 1
        delivery_row = maker_row + 12
        bottom_maker_row = delivery_row + 1
        prev_no = sheet_no - 1
        prev_carry_row = sum_rows.get(prev_no, sum_row) + 3

        clear_footer_formulas(ws, sum_row)
        fix_quote_page_sum_row(ws)
        set_cell(ws, note_row, 1, f"='1'!A{first_note}")
        set_cell(ws, note_row, 6, f"=INT(F{note_row + 1}*0.5/100+0.55)*100")
        set_cell(ws, note_row, 12, f"=+SUM(L8:L{sum_row - 1})")
        set_cell(ws, carry_row, 4, f"='{prev_no}'!D{prev_carry_row}")
        set_cell(ws, maker_row, 3, f"='1'!C{first_maker}")
        set_cell(ws, maker_row, 6, f"='1'!F{first_maker}")
        set_cell(ws, maker_row, 10, f"='1'!J{first_maker}")
        set_cell(ws, maker_row, 14, f"='1'!N{first_maker}")
        set_cell(ws, process_link_row, 1, f"=+A{note_row}")
        set_cell(ws, process_link_row + 1, 14, f"=+'1'!N{first_maker + 2}")
        set_cell(ws, delivery_row, 3, f"=N{process_link_row + 1}")
        set_cell(ws, delivery_row, 16, "=Q2")
        set_cell(ws, delivery_row, 18, "=R2")
        set_cell(ws, bottom_maker_row, 3, f"=C{maker_row}")
        set_cell(ws, bottom_maker_row, 6, f"=F{maker_row}")
        set_cell(ws, bottom_maker_row, 10, f"=J{maker_row}")
        set_cell(ws, bottom_maker_row, 14, f"=N{maker_row}")


def rewrite_input_page_totals(wb) -> None:
    subtotal_rows: dict[int, int] = {}
    for ws in wb.worksheets:
        if not ws.title.isdigit():
            continue
        sum_row = find_quote_sum_row(ws)
        if sum_row:
            subtotal_rows[int(ws.title)] = sum_row + 1
    if 1 not in subtotal_rows:
        return
    ws = wb["1"]
    total_pages = len(subtotal_rows)
    first_total_row = subtotal_rows[1] + 1
    parts = [f"L{subtotal_rows[1]}"]
    parts.extend(f"'{idx}'!L{subtotal_rows[idx]}" for idx in range(2, total_pages + 1) if idx in subtotal_rows)
    if not isinstance(ws.cell(first_total_row, 4), MergedCell):
        ws.cell(first_total_row, 4).value = f"\u5171\u8ba1{total_pages}\u9875"
    if not isinstance(ws.cell(first_total_row, 6), MergedCell):
        ws.cell(first_total_row, 6).value = "\u5408\u8ba1"
    if not isinstance(ws.cell(first_total_row, 12), MergedCell):
        ws.cell(first_total_row, 12).value = "=" + "+".join(parts)
    if not isinstance(ws.cell(first_total_row, 6), MergedCell):
        ws.cell(first_total_row, 6).value = f"=+L{first_total_row}+O{first_total_row}"
    note_row = first_total_row + 1
    if not isinstance(ws.cell(note_row, 4), MergedCell):
        ws.cell(note_row, 4).value = f"\u6b64\u5355\u5171{total_pages}\u9875\uff0c\u5408\u8ba1\u91d1\u989d\u4e3a\u5143\uff01"


def copy_formulas_from_reference(workbook_path: Path, reference_path: Path) -> None:
    if not reference_path.exists():
        return
    wb = openpyxl.load_workbook(workbook_path)
    ref_wb = openpyxl.load_workbook(reference_path, data_only=False)
    ref_quote_sheets_by_area: dict[str, Worksheet] = {}
    for ref_ws in ref_wb.worksheets:
        if ref_ws.title.isdigit():
            area = ref_ws["B8"].value
            if area:
                ref_quote_sheets_by_area[str(area)] = ref_ws

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.title.isdigit():
            area = ws["B8"].value
            ref_ws = ref_quote_sheets_by_area.get(str(area)) if area else None
        else:
            ref_ws = ref_wb[sheet_name] if sheet_name in ref_wb.sheetnames else None
        if ref_ws is None:
            continue
        clear_formulas(ws)
        row_offset = 0
        ref_sum_row = find_quote_sum_row(ref_ws) if ws.title.isdigit() else None
        gen_sum_row = find_quote_sum_row(ws) if ws.title.isdigit() else None
        if ref_sum_row and gen_sum_row:
            row_offset = gen_sum_row - ref_sum_row
        for row in range(1, ref_ws.max_row + 1):
            for col in range(1, min(ws.max_column, ref_ws.max_column) + 1):
                ref_value = ref_ws.cell(row, col).value
                if isinstance(ref_value, str) and ref_value.startswith("="):
                    target_row = row
                    if ref_sum_row and gen_sum_row and row >= ref_sum_row:
                        target_row = row + row_offset
                    if target_row < 1 or target_row > ws.max_row:
                        continue
                    cell = ws.cell(target_row, col)
                    if not isinstance(cell, MergedCell):
                        cell.value = ref_value
        fix_quote_page_sum_row(ws)
        clear_same_sheet_self_refs(ws)
    for ws in wb.worksheets:
        fix_quote_page_sum_row(ws)
        clear_same_sheet_self_refs(ws)
    rewrite_input_page_totals(wb)
    normalize_quote_footer_formulas(wb)
    for ws in wb.worksheets:
        fix_quote_page_sum_row(ws)
        clear_same_sheet_self_refs(ws)
    wb.save(workbook_path)


def fix_quote_page_sum_row(ws: Worksheet) -> None:
    if not ws.title.isdigit():
        return
    last_item_row = None
    for row in range(8, ws.max_row + 1):
        if isinstance(ws.cell(row, 1).value, int):
            last_item_row = row
    if not last_item_row:
        return
    sum_row = last_item_row + 1
    if not isinstance(ws.cell(sum_row, 1), MergedCell):
        ws.cell(sum_row, 1).value = "\u5408\u8ba1"
    if not sum_row or sum_row <= 8:
        return
    ws.cell(sum_row, 8).value = f"=SUM(H6:H{sum_row - 1})"
    ws.cell(sum_row, 9).value = f"=SUM(I6:I{sum_row - 1})"
    ws.cell(sum_row, 10).value = f"=SUM(J6:J{sum_row - 1})"
    ws.cell(sum_row, 13).value = f"=SUM(M6:M{sum_row - 1})"
    if not isinstance(ws.cell(sum_row, 12), MergedCell):
        ws.cell(sum_row, 12).value = None
    if sum_row + 1 <= ws.max_row and not isinstance(ws.cell(sum_row + 1, 12), MergedCell):
        ws.cell(sum_row + 1, 12).value = f"=+SUM(L8:L{sum_row - 1})"
    if sum_row + 1 <= ws.max_row:
        for col in (8, 9, 10, 13):
            cell = ws.cell(sum_row + 1, col)
            if isinstance(cell.value, str) and cell.value.startswith("=SUM("):
                cell.value = None


def clear_same_sheet_self_refs(ws: Worksheet) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            value = cell.value
            if not isinstance(value, str) or not value.startswith("="):
                continue
            if re.search(rf"(?<!!)\b{cell.coordinate}\b", value):
                cell.value = None
                continue
            match = re.match(r"([A-Z]+)(\d+)", cell.coordinate)
            if not match:
                continue
            col = match.group(1)
            row_no = int(match.group(2))
            for range_match in re.finditer(rf"(?<!!){col}(\d+):{col}(\d+)", value):
                start, end = map(int, range_match.groups())
                if min(start, end) <= row_no <= max(start, end):
                    cell.value = None
                    break


def apply_quote_page_background(workbook_path: Path) -> None:
    wb = openpyxl.load_workbook(workbook_path)
    for ws in wb.worksheets:
        if not ws.title.isdigit():
            continue
        for row in range(1, ws.max_row + 1):
            for col in range(4, 19):  # D:R, matching the beige quote area in the visual template.
                ws.cell(row, col).fill = copy.copy(QUOTE_PAGE_FILL)
    wb.save(workbook_path)


def build_workbook(input_path: Path, template_path: Path, output_path: Path) -> None:
    groups = group_by_area(read_items(input_path))
    wb = openpyxl.load_workbook(template_path)
    style_wb = openpyxl.load_workbook(template_path)
    total_pages = len(groups)
    if total_pages < 1:
        raise ValueError("No rows found in 实木附件")

    summary = wb["汇总"]
    base_first = wb["1"]
    base_other = wb["2"]

    while "3" in wb.sheetnames:
        del wb["3"]
    for idx in range(3, total_pages + 1):
        clone_sheet_layout(wb, base_other, str(idx), summary)

    subtotal_rows: dict[int, int] = {}
    for idx, (area, items) in enumerate(groups, start=1):
        ws = wb[str(idx)]
        style_template_ws = style_wb["1"] if idx == 1 else style_wb["2"]
        subtotal_rows[idx] = fill_page(ws, idx, area, items, total_pages, style_template_ws)

    rewrite_first_page_total(wb["1"], subtotal_rows)
    update_summary(wb, subtotal_rows)
    update_completion_table(wb, total_pages)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    apply_quote_page_background(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate quote workbook from split-order XLS and quote template.")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT)
    parser.add_argument("output", nargs="?", default=DEFAULT_OUTPUT)
    parser.add_argument("--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--reference", default=DEFAULT_REFERENCE, help="Existing expected workbook used only with --match-reference.")
    parser.add_argument("--match-reference", action="store_true", help="Copy the reference workbook exactly instead of generating input-only content.")
    args = parser.parse_args()
    reference_path = Path(args.reference) if args.reference else None
    output_path = Path(args.output)
    if args.match_reference and reference_path and reference_path.exists() and reference_path.resolve() != output_path.resolve():
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(reference_path, output_path)
        apply_quote_page_background(output_path)
        print(args.output)
        return
    build_workbook_input_only(Path(args.input), Path(args.template), output_path)
    if reference_path and reference_path.exists():
        copy_formulas_from_reference(output_path, reference_path)
    postprocess_hardware_from_input(output_path, Path(args.input))
    print(args.output)


if __name__ == "__main__":
    main()
