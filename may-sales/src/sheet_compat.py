# -*- coding: utf-8 -*-
"""把新版三件套（2026-10 起的导出）转成 generate_may_sales_report.js 认的老版式。

老脚本按**固定列号**取值（联思取第 8/14/20 列、奢匠取第 62/64 列），
导出格式一变就静默算错或直接报 Sheet 找不到。2026-10 的导出漂了三处：

1. 联思：sheet 名由 `Sheet` 改成「经销商财务统计」，前面多了 3 行标题；
   列也重排——订单总金额→原价总金额、折后金额→应收金额、
   订单类型从第 20 列挪到第 22 列（中间插了折扣名称/折后柜身/折后五金/汇款比例）。
2. 奢匠下单表：表头行由第 3 行挪到第 1 行，列数 64→65（多了「优惠金额」），
   总金额、实际销售额各后移一列。
3. 综合查询：sheet 名变成「综合查询 (原)」，列没变。

所以这里按**表头名**重新映射成老版式，列顺序再变也不会算错。
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

# worker 镜像里没有 openpyxl，服务目录里带了一份（_vendor），优先用它
_VENDOR = Path(__file__).resolve().parent / "_vendor"
if _VENDOR.is_dir() and str(_VENDOR) not in sys.path:
    sys.path.insert(0, str(_VENDOR))

import openpyxl  # noqa: E402

SCAN_ROWS = 12

LEGACY_LIANSI_HEADER = [
    "财务审核时间", "合同号", "经销商名称", "省", "市", "门店类型", "客户",
    "订单总金额", "加减收金额", "柜身金额", "五金金额", "折扣", "折扣金额",
    "折后金额", "已付金额", "未付金额", "经销商打款金额", "打款账号",
    "打款备注", "订单类型",
]

# 老列名 → 新导出里可能的写法
LIANSI_ALIASES = {
    "订单总金额": ("订单总金额", "原价总金额", "订单金额"),
    "折扣": ("折扣", "折扣名称"),
    "折后金额": ("折后金额", "应收金额"),
    "已付金额": ("已付金额", "已收款金额"),
}

# 奢匠下单表在老脚本里认的列号
SHEJIANG_INDEX = {"月份": 1, "省份": 3, "地区": 4, "经销商": 5, "订单类型": 6,
                  "总金额": 61, "实际销售额": 63}


def _cells(row):
    out = []
    for c in row:
        if c is None:
            out.append("")
        elif isinstance(c, float) and c.is_integer():
            out.append(str(int(c)))
        else:
            out.append(str(c).strip())
    return out


def _find_header(rows, *keywords):
    """找表头行：前 12 行里第一行同时含全部关键词的。"""
    for i, row in enumerate(rows[:SCAN_ROWS]):
        cells = _cells(row)
        joined = "|".join(cells)
        if all(k in joined for k in keywords):
            return i
    return -1


def _col(header_cells, *aliases):
    for a in aliases:
        for i, c in enumerate(header_cells):
            if c == a:
                return i
    for a in aliases:
        for i, c in enumerate(header_cells):
            if a and a in c:
                return i
    return -1


def _out_path(path: Path, tag: str) -> Path:
    return path.with_name(f"_norm_{tag}_{path.name}")


def _save(wb, path: Path) -> Path:
    wb.save(path)
    return path


def normalize_zhcx(path: str) -> str:
    """综合查询：只把 sheet 名改回「综合查询」，列不动。"""
    p = Path(path)
    wb = openpyxl.load_workbook(p, data_only=True)
    if "综合查询" in wb.sheetnames:
        return str(p)
    target = next((n for n in wb.sheetnames if "综合查询" in n), wb.sheetnames[0])
    wb[target].title = "综合查询"
    return str(_save(wb, _out_path(p, "zhcx")))


def normalize_liansi(path: str) -> str:
    """联思：按表头名重排成老的 20 列，sheet 命名回 Sheet。"""
    p = Path(path)
    wb = openpyxl.load_workbook(p, data_only=True)
    if "Sheet" in wb.sheetnames:
        return str(p)

    best, best_rows, best_hdr = None, None, -1
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        hdr = _find_header(rows, "订单类型", "经销商")
        if hdr >= 0 and _col(_cells(rows[hdr]), "财务审核时间", "合同号") >= 0:
            best, best_rows, best_hdr = ws, rows, hdr
            break
    if best is None:
        return str(p)          # 认不出就原样交给下游，让老脚本自己报错

    hdr = best_hdr
    header_cells = _cells(best_rows[hdr])
    idx = {}
    for name in LEGACY_LIANSI_HEADER:
        aliases = LIANSI_ALIASES.get(name, (name,))
        idx[name] = _col(header_cells, *aliases)

    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Sheet"
    ws.append(LEGACY_LIANSI_HEADER)
    for row in best_rows[hdr + 1:]:
        if not any(v not in (None, "") for v in row):
            continue
        ws.append([row[idx[n]] if idx[n] >= 0 and idx[n] < len(row) else None
                   for n in LEGACY_LIANSI_HEADER])
    return str(_save(out, _out_path(p, "liansi")))


def normalize_shejiang(path: str) -> str:
    """奢匠下单表：表头行拉回第 3 行、金额列按列名摆回老位置。

    另外两个 Sheet 不能丢：销量情况报表还要读「直营店+电商下单表」，
    所以这里是「复制整本 + 换掉奢匠下单表」，不是只留一个 Sheet。
    """
    p = Path(path)
    wb = openpyxl.load_workbook(p, data_only=True)
    name = next((n for n in wb.sheetnames if n.strip() == "奢匠下单表"), None)
    if name is None:
        name = next((n for n in wb.sheetnames if "下单" in n and "直营" not in n), None)
    if name is None:
        return str(p)

    rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
    hdr = _find_header(rows, "月份", "经销商")
    if hdr < 0:
        return str(p)
    if hdr == 2 and _col(_cells(rows[2]), "总金额") == 61:
        return str(p)          # 已经是老版式

    header_cells = _cells(rows[hdr])
    idx = {k: _col(header_cells, k) for k in SHEJIANG_INDEX}
    if idx["总金额"] < 0 or idx["实际销售额"] < 0 or idx["经销商"] < 0:
        return str(p)

    width = 64
    out = openpyxl.Workbook()
    out.remove(out.active)
    # 先把其它 Sheet 原样搬过去（销量情况报表要用直营店+电商下单表）
    for ws in wb.worksheets:
        if ws.title == name:
            continue
        copy = out.create_sheet(ws.title[:31])
        for row in ws.iter_rows(values_only=True):
            copy.append(list(row))
    ws = out.create_sheet("奢匠下单表")
    ws.append(["2026奢匠下单统计（已归一版式）"])
    ws.append(["由 sheet_compat 生成，仅供老脚本按列号读取"])
    head = [""] * width
    for k, pos in SHEJIANG_INDEX.items():
        head[pos] = k
    head[22] = "总金额"        # 销量情况报表按 c22 取总金额
    ws.append(head)
    for row in rows[hdr + 1:]:
        if not any(v not in (None, "") for v in row):
            continue
        line = [None] * width
        for k, pos in SHEJIANG_INDEX.items():
            i = idx[k]
            if i >= 0 and i < len(row):
                line[pos] = row[i]
        line[22] = line[61]    # 同一份总金额，两个脚本各看各的列
        ws.append(line)
    return str(_save(out, _out_path(p, "shejiang")))


NORMALIZERS = {"zhcx": normalize_zhcx, "liansi": normalize_liansi,
               "shejiang": normalize_shejiang}


def normalize_all(file_map: dict) -> dict:
    """就地归一 file_map（{'zhcx'/'liansi'/'shejiang': 路径}），返回新 map。"""
    out = dict(file_map)
    for kind, fn in NORMALIZERS.items():
        path = out.get(kind)
        if not path:
            continue
        try:
            out[kind] = fn(path)
        except Exception:  # noqa: BLE001
            # 归一失败就把原文件交下去，宁可报老错误也别吞掉
            pass
    return out


if __name__ == "__main__":       # 手工验证：python sheet_compat.py <kind> <file>
    import sys
    print(NORMALIZERS[sys.argv[1]](sys.argv[2]))
