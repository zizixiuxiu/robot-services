"""下车间明细表结构冲突检测（只读，不修改任何单元格）。

用途
----
客户上传的原始报价单是人工维护的，列位置、合并区域经常和生成器预期不一致。
生成器为了补「工分」等内容，会往固定列写值——一旦那个位置被别的合并块
占着，写进去就会**覆盖原件内容**（真实案例：颜色文字被工分公式顶掉）。

本模块在动手之前把这类冲突全部列出来，供人工确认，避免静默改坏原表。

检测的是「结构冲突」：合并区域与关键列位置不一致、关键列被合并覆盖、
关键列表头缺失、合并跨度与产品行范围不符。

用法
----
    python conflict_check.py <workbook.xlsx> [--json out.json]

作为库调用：

    from conflict_check import check_workbook
    conflicts = check_workbook(path)      # -> list[Conflict]

每个 Conflict 都有 sheet / location / code / summary / detail / impact 字段。
"""
from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None

# 表头识别：这些文本用于定位列
HEADER_KEYWORDS = {
    "序号": ("序号",),
    "产品名称": ("产品名称",),
    "小计": ("小计",),
    "工分": ("工分",),
    "颜色": ("颜色",),
    "材质": ("材质",),
}

MAX_HEADER_SCAN_ROWS = 20


@dataclass
class Conflict:
    sheet: str
    code: str          # C1..C6
    location: str      # 例如 "M8:O24"
    summary: str       # 一句话说明
    detail: str        # 原件现状
    impact: str        # 会导致什么

    def as_row(self) -> str:
        return f"[{self.code}] {self.sheet} {self.location}｜{self.summary}"


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\u3000", " ").strip()


def _find_header_row(ws: Any) -> int | None:
    """找同时包含「序号」和「产品名称」的表头行。"""
    for row_idx in range(1, min(ws.max_row, MAX_HEADER_SCAN_ROWS) + 1):
        texts = {_text(ws.cell(row_idx, c).value) for c in range(1, min(ws.max_column, 40) + 1)}
        if any(k in t for t in texts for k in HEADER_KEYWORDS["序号"]) and \
           any(k in t for t in texts for k in HEADER_KEYWORDS["产品名称"]):
            return row_idx
    return None


def _locate_columns(ws: Any, header_row: int) -> dict[str, int]:
    """返回 {逻辑列名: 列号}，找不到的不出现在结果里。"""
    found: dict[str, int] = {}
    for col in range(1, min(ws.max_column, 40) + 1):
        raw = _text(ws.cell(header_row, col).value)
        if not raw:
            continue
        for logical, keys in HEADER_KEYWORDS.items():
            if logical in found:
                continue
            if any(k in raw for k in keys):
                found[logical] = col
    return found


def _product_merges(ws: Any, start_row: int, end_row: int) -> list[Any]:
    """只看产品行范围内的纵向合并块。

    必须同时满足：
      - 与 [start_row, end_row] 有交集（排除合计行/备注行/说明行的合并）
      - 纵向跨至少两行（排除「一级部门」这类单行横向标签合并）
    """
    out = []
    for rng in ws.merged_cells.ranges:
        if rng.max_row < start_row or rng.min_row > end_row:
            continue
        if rng.max_row <= rng.min_row:
            continue
        out.append(rng)
    return out


def _last_product_row(ws: Any, header_row: int, product_col: int, subtotal_col: int | None) -> int:
    last = header_row
    for row_idx in range(header_row + 1, ws.max_row + 1):
        label = _text(ws.cell(row_idx, 1).value)
        if label.startswith("合计") or label.startswith("预付"):
            break
        if _text(ws.cell(row_idx, product_col).value):
            last = row_idx
        elif subtotal_col and ws.cell(row_idx, subtotal_col).value not in (None, ""):
            last = row_idx
    return last


def _check_sheet(ws: Any) -> list[Conflict]:
    out: list[Conflict] = []
    header_row = _find_header_row(ws)
    if header_row is None:
        return out

    cols = _locate_columns(ws, header_row)
    start_row = header_row + 1
    sheet = ws.title

    prod_col = cols.get("产品名称")
    sub_col = cols.get("小计")
    wp_col = cols.get("工分")
    color_col = cols.get("颜色")
    mat_col = cols.get("材质")

    if prod_col is None:
        return out

    end_row = _last_product_row(ws, header_row, prod_col, sub_col)
    if end_row < start_row:
        return out

    merges = _product_merges(ws, start_row, end_row)

    # C4：工分列表头缺失
    if wp_col is None:
        out.append(Conflict(
            sheet, "C4", f"第{header_row}行",
            "未找到「工分」列表头",
            "表头行里没有工分列",
            "生成器不会补填工分，工分整列为空",
        ))

    handled: set[str] = set()

    for rng in merges:
        span = str(rng)
        spans_columns = rng.max_col > rng.min_col  # 单列纵向合并（如材质跨几行）属正常
        covers_color = color_col is not None and rng.min_col <= color_col <= rng.max_col
        covers_wp = wp_col is not None and rng.min_col <= wp_col <= rng.max_col
        covers_mat = mat_col is not None and rng.min_col <= mat_col <= rng.max_col

        # C1：颜色块起点偏左，把左边的列（通常是工分）一起吞了
        if covers_color and color_col is not None and rng.min_col < color_col:
            swallowed = [n for n, c in (("工分", wp_col),) if c is not None and rng.min_col <= c < color_col]
            extra = f"，并把「{swallowed[0]}」列一起并入" if swallowed else ""
            out.append(Conflict(
                sheet, "C1", span,
                f"颜色块合并起点比颜色表头列靠左{extra}",
                f"颜色表头在第 {color_col} 列，但数据区合并块从第 {rng.min_col} 列开始（{span}）",
                "颜色说明被挪进左侧列存储；生成器往那一列写入工分公式时，"
                "会把颜色文字覆盖掉，颜色列则显示成工分数字",
            ))
            handled.add(span)
            # C5：跨度与产品行不符
            if abs(rng.max_row - end_row) > 1:
                out.append(Conflict(
                    sheet, "C5", span,
                    "颜色块合并跨度与产品行范围不一致",
                    f"合并到第 {rng.max_row} 行，最后一个产品行是第 {end_row} 行",
                    "颜色说明会溢出到空白行，或末尾几行没有颜色说明",
                ))
            continue

        # C2：工分列被跨列合并块覆盖（但没覆盖颜色列）
        if covers_wp and not covers_color and spans_columns:
            out.append(Conflict(
                sheet, "C2", span,
                "「工分」列被跨列合并块覆盖",
                f"合并区域 {span} 横跨第 {rng.min_col}~{rng.max_col} 列，包含工分列（第 {wp_col} 列）",
                "工分无法按行填写；写入只能落在合并块左上角，会覆盖该格原有内容",
            ))
            handled.add(span)

        # C6：材质列被跨列合并块覆盖（且与颜色、工分无关）
        if covers_mat and not covers_color and not covers_wp and spans_columns:
            out.append(Conflict(
                sheet, "C6", span,
                "「材质」列被跨列合并块覆盖",
                f"合并区域 {span} 包含材质列（第 {mat_col} 列）",
                "材质明细可能显示不全或无法逐行填写",
            ))

    return out


def check_workbook(path: str | Path) -> list[Conflict]:
    """只读检查整个工作簿，返回全部冲突（不改动文件）。"""
    if openpyxl is None:
        raise RuntimeError("需要 openpyxl")
    wb = openpyxl.load_workbook(str(path))
    conflicts: list[Conflict] = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        conflicts.extend(_check_sheet(ws))
    return conflicts


def summarize(conflicts: list[Conflict]) -> str:
    if not conflicts:
        return "未发现结构冲突。"
    by_sheet: dict[str, int] = {}
    for c in conflicts:
        by_sheet[c.sheet] = by_sheet.get(c.sheet, 0) + 1
    lines = [f"共发现 {len(conflicts)} 处冲突，涉及 {len(by_sheet)} 个工作表："]
    for sheet, n in by_sheet.items():
        lines.append(f"  - 工作表 {sheet}：{n} 处")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="下车间明细表结构冲突检测（只读）")
    parser.add_argument("workbook")
    parser.add_argument("--json", help="把结果另存为 JSON")
    args = parser.parse_args(argv)

    path = Path(args.workbook)
    if not path.exists():
        print(f"文件不存在：{path}")
        return 2

    conflicts = check_workbook(path)
    print(f"检查文件：{path}")
    print(summarize(conflicts))
    if conflicts:
        print()
        for c in conflicts:
            print(f"[{c.code}] 工作表 {c.sheet}  位置 {c.location}")
            print(f"      问题：{c.summary}")
            print(f"      现状：{c.detail}")
            print(f"      影响：{c.impact}")
            print()

    if args.json:
        Path(args.json).write_text(
            json.dumps([asdict(c) for c in conflicts], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"已保存 JSON：{args.json}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
