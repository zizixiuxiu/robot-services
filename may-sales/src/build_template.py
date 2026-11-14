# -*- coding: utf-8 -*-
"""用最新的经销商名单重建「销售部业绩核对表」模板。

老模板写死了 113 家（2026-05 名单）和合计行 116，名单一变就对不上：
新商落在模板外被丢掉、合计行还固定引用 3:115。这里拿老模板当格式母版，
只换名单区行数，合计行和公式按实际行数重算。

用法：
  python build_template.py <名单xlsx> <名单sheet> <老模板> <新模板>
名单表头需含：序号/区域/省份/城市/经销商（按表头名找列，不靠列号）
"""
from __future__ import annotations

import sys
from copy import copy
from pathlib import Path

_VENDOR = Path(__file__).resolve().parent / "_vendor"
if _VENDOR.is_dir() and str(_VENDOR) not in sys.path:
    sys.path.insert(0, str(_VENDOR))

import openpyxl
from openpyxl.utils import get_column_letter

HEADER_ROW = 2
FIRST_DATA_ROW = HEADER_ROW + 1
COLS = {"seq": 1, "region": 2, "province": 3, "city": 4, "dealer": 5}
SUM_COLS = (6, 7, 9, 10, 12, 13)      # F G I J L M
CLEAR_COLS = (6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)


def read_dealers(path: str, sheet_name: str) -> list[tuple]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        raise SystemExit(f"名单文件是空的: {path}")

    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    missing = [n for n in ("区域", "省份", "城市", "经销商") if n not in header]
    if missing:
        raise SystemExit(f"名单表头缺少 {missing}，实际表头: {header}")
    idx = {n: header.index(n) for n in ("序号", "区域", "省份", "城市", "经销商")}

    out = []
    for r in rows[1:]:
        dealer = str(r[idx["经销商"]] or "").strip() if idx["经销商"] < len(r) else ""
        if not dealer or dealer == "合计":
            continue
        seq = r[idx["序号"]] if idx["序号"] >= 0 and idx["序号"] < len(r) else None
        out.append((
            seq if seq not in (None, "") else len(out) + 1,
            str(r[idx["区域"]] or "").strip(),
            str(r[idx["省份"]] or "").strip(),
            str(r[idx["城市"]] or "").strip(),
            dealer,
        ))
    return out


def find_total_row(ws) -> int:
    for r in range(ws.max_row, 1, -1):
        if str(ws.cell(row=r, column=1).value or "").strip() == "合计":
            return r
    raise SystemExit(f"sheet {ws.title} 里找不到「合计」行")


def _resize(ws, new_n: int, old_n: int, total: int) -> None:
    """把数据区行数从 old_n 调成 new_n（合计行随之移动）。"""
    if new_n < old_n:
        ws.delete_rows(FIRST_DATA_ROW + new_n, old_n - new_n)
    elif new_n > old_n:
        ws.insert_rows(total, new_n - old_n)
        # 新插入的行没有样式，复制最后一条数据行的样式
        for r in range(total, total + new_n - old_n):
            for c in range(1, ws.max_column + 1):
                src = ws.cell(row=total - 1, column=c)
                dst = ws.cell(row=r, column=c)
                if src.has_style:
                    dst._style = copy(src._style)


def rebuild(ws, dealers: list[tuple]) -> tuple[int, int]:
    total = find_total_row(ws)
    old_n = total - FIRST_DATA_ROW
    new_n = len(dealers)

    # 合并区先拆掉：delete_rows 会移动/丢弃它们，等行数定了再按新位置合并
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    _resize(ws, new_n, old_n, total)
    total = FIRST_DATA_ROW + new_n

    # 写名单
    for i, (seq, region, prov, city, dealer) in enumerate(dealers):
        r = FIRST_DATA_ROW + i
        for c, v in ((1, seq), (2, region), (3, prov), (4, city), (5, dealer)):
            ws.cell(row=r, column=c).value = v
        # 注意：ws.cell(..., value=None) 不会真的清值（openpyxl 只在 value 非 None 时赋值）
        for c in CLEAR_COLS:
            ws.cell(row=r, column=c).value = None
        ws.cell(row=r, column=8).value = f"=SUM(F{r}:G{r})"
        ws.cell(row=r, column=14).value = f"=SUM(L{r}:M{r})"

    # 合计行
    ws.cell(row=total, column=1).value = "合计"
    for c in range(2, 6):
        ws.cell(row=total, column=c).value = None
    for c in SUM_COLS:
        letter = get_column_letter(c)
        ws.cell(row=total, column=c).value = f"=SUM({letter}{FIRST_DATA_ROW}:{letter}{total - 1})"
    ws.cell(row=total, column=8).value = f"=SUM(F{total}:G{total})"
    ws.cell(row=total, column=14).value = f"=SUM(L{total}:M{total})"
    for c in (11, 15, 16):
        ws.cell(row=total, column=c).value = None

    # 删掉尾部可能残留的空行（delete_rows 只清单元格，行高/格式条目会留下空 <row>）
    if ws.max_row > total:
        ws.delete_rows(total + 1, ws.max_row - total)
    for r in [k for k in list(ws.row_dimensions.keys()) if k > total]:
        del ws.row_dimensions[r]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    ws.merge_cells(start_row=total, start_column=1, end_row=total, end_column=5)
    return old_n, total


def main() -> None:
    if len(sys.argv) < 5:
        raise SystemExit("用法: python build_template.py <名单xlsx> <名单sheet> <老模板> <新模板>")
    list_file, list_sheet, tpl, out = sys.argv[1:5]

    dealers = read_dealers(list_file, list_sheet)
    print(f"名单共 {len(dealers)} 家: " + "、".join(d[4] for d in dealers[:3]) + " ...")

    wb = openpyxl.load_workbook(tpl)
    for name in ("Sheet1", "Sheet1 (2)"):
        if name not in wb.sheetnames:
            raise SystemExit(f"模板缺少 sheet: {name}")
        old_n, total = rebuild(wb[name], dealers)
        print(f"{name}: 数据行 {old_n} → {len(dealers)}，合计行第 {total} 行")

    # 标题里的月份跟着名单走（名单 sheet 名带「10月」时）
    import re
    m = re.search(r"(\d+)\s*月", list_sheet or "")
    if m:
        month = int(m.group(1))
        for name in ("Sheet1", "Sheet1 (2)"):
            cell = wb[name].cell(row=1, column=1)
            if isinstance(cell.value, str):
                cell.value = re.sub(r"\d+\s*月", f"{month}月", cell.value)

    wb.save(out)
    print(f"已生成新模板: {out}")


if __name__ == "__main__":
    main()
