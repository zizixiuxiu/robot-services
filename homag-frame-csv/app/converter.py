from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

from app.color import convert as convert_color
from app.color import split_batch_and_color


OUTPUT_HEADERS = [
    "订单号", "板件名称", "加工长度", "加工宽度", "数量", "材料描述", "纹理",
    "完工长度", "完工宽度", "打孔工艺", "特殊工艺", "正面条码", "开槽工艺",
    "客户", "柜体名称", "分柜号", "品牌", "厚度", "面积", "批次号", "工艺路线",
    "材料描述2", "物料编码", "异型", "订单类型", "封边类型",
]

SOURCE_SHEETS = ("门套", "窗套")
REQUIRED_HEADERS = ("大类", "开料长", "开料宽", "数量", "厚度", "单双", "颜色", "客户名称", "材质")
CSV_SHEET_NAME = "门套"


class ConversionError(ValueError):
    """输入表格不符合转换要求。"""


@dataclass(frozen=True)
class ConversionStats:
    door_rows: int
    window_rows: int
    output_rows: int
    quantity_sum: float
    skipped_empty: int
    huqiang_rows: int
    discarded_shoucai: int
    color_no_slash: int
    missing_sheets: tuple[str, ...]
    sheet_names: tuple[str, ...]
    has_huqiang_file: bool


@dataclass(frozen=True)
class ConversionOutput:
    csv_bytes: bytes
    huqiang_bytes: bytes | None
    stats: ConversionStats


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_number(value: Any) -> int | float | str:
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = clean_text(value)
        if not text:
            return ""
        try:
            number = float(text)
        except ValueError:
            return text
    return int(number) if number.is_integer() else number


def _header_index(headers: Sequence[Any]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for index, raw in enumerate(headers):
        name = clean_text(raw)
        if name and name not in mapping:
            mapping[name] = index
    missing = [name for name in REQUIRED_HEADERS if name not in mapping]
    if missing:
        raise ConversionError("源表缺少字段：" + "、".join(missing))
    return mapping


def _cell(row: Sequence[Any], index_map: dict[str, int], name: str) -> Any:
    index = index_map.get(name)
    if index is None or index >= len(row):
        return ""
    return row[index]


def _join_special(face: str, material: str) -> str:
    parts = [part for part in (clean_text(face), clean_text(material)) if part]
    return " ".join(parts)


def transform_rows(
    headers: Sequence[Any], source_rows: Sequence[Sequence[Any]]
) -> tuple[list[list[Any]], int, int, float]:
    index_map = _header_index(headers)
    output_rows: list[list[Any]] = []
    skipped_empty = 0
    color_no_slash = 0
    quantity_sum = 0.0

    for row in source_rows:
        category = clean_text(_cell(row, index_map, "大类"))
        color_raw = clean_text(_cell(row, index_map, "颜色"))
        if not category or not color_raw:
            skipped_empty += 1
            continue

        batch, color_code = split_batch_and_color(color_raw)
        if "/" not in color_raw:
            color_no_slash += 1
        ysm, original = convert_color(color_code)
        qty = clean_number(_cell(row, index_map, "数量"))
        if isinstance(qty, (int, float)):
            quantity_sum += float(qty)

        length = clean_number(_cell(row, index_map, "开料长"))
        width = clean_number(_cell(row, index_map, "开料宽"))
        out = [""] * len(OUTPUT_HEADERS)
        out[0] = batch
        out[1] = category
        out[2] = length
        out[3] = width
        out[4] = qty
        out[5] = ysm
        out[6] = 1
        out[7] = length
        out[8] = width
        out[10] = _join_special(
            _cell(row, index_map, "单双"),
            _cell(row, index_map, "材质"),
        )
        out[13] = clean_text(_cell(row, index_map, "客户名称"))
        out[17] = clean_number(_cell(row, index_map, "厚度"))
        out[21] = original
        output_rows.append(out)

    return output_rows, skipped_empty, color_no_slash, quantity_sum


def _first_nonempty_row(values: list[list[Any]]) -> int | None:
    return next(
        (index for index, row in enumerate(values) if any(clean_text(value) for value in row)),
        None,
    )


def _sheet_values_from_rows(values: list[list[Any]]) -> tuple[list[Any], list[list[Any]]]:
    first = _first_nonempty_row(values)
    if first is None:
        raise ConversionError("源表为空")
    return values[first], values[first + 1 :]


def read_sheets(content: bytes, filename: str) -> dict[str, tuple[list[Any], list[list[Any]]]]:
    suffix = Path(filename).suffix.lower()
    sheets: dict[str, list[list[Any]]] = {}

    if suffix == ".csv":
        for encoding in ("gb18030", "utf-8-sig"):
            try:
                text = content.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ConversionError("无法读取 .csv 文件：编码不是 GB18030 或 UTF-8")
        sheets[CSV_SHEET_NAME] = [row for row in csv.reader(io.StringIO(text))]
    elif suffix == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise RuntimeError("缺少 xlrd 依赖，无法读取 .xls") from exc
        try:
            workbook = xlrd.open_workbook(file_contents=content)
            for name in workbook.sheet_names():
                sheet = workbook.sheet_by_name(name)
                sheets[clean_text(name)] = [
                    sheet.row_values(index) for index in range(sheet.nrows)
                ]
        except ConversionError:
            raise
        except Exception as exc:
            raise ConversionError(f"无法读取 .xls 文件：{exc}") from exc
    elif suffix == ".xlsx":
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise RuntimeError("缺少 openpyxl 依赖，无法读取 .xlsx") from exc
        try:
            workbook = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
            for sheet in workbook.worksheets:
                sheets[clean_text(sheet.title)] = [list(row) for row in sheet.iter_rows(values_only=True)]
        except ConversionError:
            raise
        except Exception as exc:
            raise ConversionError(f"无法读取 .xlsx 文件：{exc}") from exc
    else:
        raise ConversionError("只支持 .xls、.xlsx 或 .csv 文件")

    parsed: dict[str, tuple[list[Any], list[list[Any]]]] = {}
    for name, values in sheets.items():
        if not values:
            continue
        parsed[name] = _sheet_values_from_rows(values)
    if not parsed:
        raise ConversionError("源表为空")
    return parsed


def write_csv(headers: Sequence[Any], rows: Sequence[Sequence[Any]]) -> bytes:
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, lineterminator="\r\n")
    writer.writerow([clean_text(h) for h in headers])
    writer.writerows(rows)
    return buffer.getvalue().encode("gb18030")


def _passthrough_cell(value: Any) -> Any:
    if value is None:
        return None
    return value


def write_xlsx_sheet(title: str, headers: Sequence[Any], rows: Sequence[Sequence[Any]]) -> bytes:
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title[:31] or "护墙"
    sheet.append(list(headers))
    for row in rows:
        sheet.append([_passthrough_cell(value) for value in row])
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _nonempty_row_count(rows: Sequence[Sequence[Any]]) -> int:
    return sum(1 for row in rows if any(clean_text(value) for value in row))


def convert_workbook(content: bytes, filename: str) -> ConversionOutput:
    sheets = read_sheets(content, filename)
    sheet_names = tuple(sheets.keys())
    suffix = Path(filename).suffix.lower()

    output_rows: list[list[Any]] = []
    door_rows = 0
    window_rows = 0
    skipped_empty = 0
    color_no_slash = 0
    quantity_sum = 0.0
    missing_sheets: list[str] = []

    if suffix == ".csv":
        headers, rows = next(iter(sheets.values()))
        converted, skipped, no_slash, qty = transform_rows(headers, rows)
        output_rows.extend(converted)
        door_rows = len(converted)
        skipped_empty += skipped
        color_no_slash += no_slash
        quantity_sum += qty
    else:
        for name in SOURCE_SHEETS:
            if name not in sheets:
                missing_sheets.append(name)
                continue
            headers, rows = sheets[name]
            converted, skipped, no_slash, qty = transform_rows(headers, rows)
            output_rows.extend(converted)
            skipped_empty += skipped
            color_no_slash += no_slash
            quantity_sum += qty
            if name == "门套":
                door_rows = len(converted)
            else:
                window_rows = len(converted)

        if "门套" not in sheets and "窗套" not in sheets:
            raise ConversionError(
                "未找到「门套」「窗套」工作表，当前表：" + "、".join(sheet_names or ["(空)"])
            )

    huqiang_bytes = None
    huqiang_rows = 0
    if "护墙" in sheets:
        headers, rows = sheets["护墙"]
        huqiang_rows = _nonempty_row_count(rows)
        if huqiang_rows:
            huqiang_bytes = write_xlsx_sheet("护墙", headers, rows)

    discarded_shoucai = 0
    if "手裁" in sheets:
        discarded_shoucai = _nonempty_row_count(sheets["手裁"][1])

    if not output_rows:
        raise ConversionError("门套/窗套没有可转换的数据行")

    stats = ConversionStats(
        door_rows=door_rows,
        window_rows=window_rows,
        output_rows=len(output_rows),
        quantity_sum=quantity_sum,
        skipped_empty=skipped_empty,
        huqiang_rows=huqiang_rows,
        discarded_shoucai=discarded_shoucai,
        color_no_slash=color_no_slash,
        missing_sheets=tuple(missing_sheets),
        sheet_names=sheet_names,
        has_huqiang_file=huqiang_bytes is not None,
    )
    return ConversionOutput(
        csv_bytes=write_csv(OUTPUT_HEADERS, output_rows),
        huqiang_bytes=huqiang_bytes,
        stats=stats,
    )


def convert_excel_to_csv(content: bytes, filename: str) -> tuple[bytes, ConversionStats]:
    result = convert_workbook(content, filename)
    return result.csv_bytes, result.stats
