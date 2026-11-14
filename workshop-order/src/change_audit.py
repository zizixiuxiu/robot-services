"""下车间改动审计：逐格比对「原表」与「生成结果」，把每一处出入分类列出来。

目的
----
生成过程必然有改动（单价打折、补工分、清木箱行、清备注里的结算金额……）。
这些改动必须能逐条说明；除此之外的单元格必须**一字未改**，并由本模块给出断言。

做法
----
1. 逐表做**行对齐**（difflib.SequenceMatcher，以整行内容为签名），
   这样「删掉一行」会报成一行删除，而不是几千条单元格位移；
2. 对替换区块再做单元格级比对；
3. 每处差异归到三类之一：
   - 形式转换：公式固化成数值，结果不变（字面不同、值相同）
   - 规则改动：代码里写明的业务规则（单价折扣 / 工分补填 / 木箱行清空 / 备注清理）
   - 需人工确认：真正的意外出入（结果变了、内容改写、未归类）
4. 「需人工确认」的条数必须为 0，否则说明有未声明的改动。

用法
----
    python change_audit.py 原表.xlsx 结果.xlsx [--json out.json] [--limit 8]
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

import openpyxl

MAX_HEADER_SCAN_ROWS = 20
SINGLE_REF_RE = re.compile(r"^=\s*\+?\s*\$?([A-Z]{1,3})\$?(\d+)$")
# 整格都是报价/结算信息（金额、折扣）——备注清理规则会整格删掉
_PAYMENT_HINT = re.compile(r"(优惠|折扣|折后|需支付|应付|金额|定金)")


def _looks_like_payment(text: str) -> bool:
    return "元" in text and bool(_PAYMENT_HINT.search(text))
HEADER_KEYS = {
    "单价": ("单价",),
    "小计": ("小计",),
    "工分": ("工分",),
    "颜色": ("颜色",),
    "备注": ("备注",),
    "产品名称": ("产品名称",),
}

# ── 归类 ────────────────────────────────────────────────────────────
CAT_FREEZE = "公式固化为数值（结果不变）"
CAT_FORMULA = "公式改写"
CAT_PRICE = "单价折扣"
CAT_WORKPOINT = "工分补填"
CAT_NOTE = "备注改写"
CAT_NOTE_CLEAN = "备注清理（删掉结算/折扣金额）"
CAT_WOODBOX = "木箱包装行清空"
CAT_PAGE_SUFFIX = "订单编号后缀更正"
CAT_DERIVED = "连带更新（跟随被引用单元格）"
CAT_VALUE_CHANGED = "公式结果变了"
CAT_ADDED = "新增内容"
CAT_DELETED = "删除内容"
CAT_CHANGED = "内容改写"
CAT_UNKNOWN = "未归类"

# 形式转换：字面变了，取值不变
GROUP_FORM = {CAT_FREEZE, CAT_FORMULA}
# 规则改动：代码里写明的业务规则
GROUP_RULE = {CAT_PRICE, CAT_WORKPOINT, CAT_NOTE, CAT_NOTE_CLEAN, CAT_WOODBOX,
              CAT_PAGE_SUFFIX, CAT_DERIVED}
# 需要人工确认
GROUP_REVIEW = {CAT_VALUE_CHANGED, CAT_ADDED, CAT_DELETED, CAT_CHANGED, CAT_UNKNOWN}

KIND_FORM = "形式转换"
KIND_RULE = "规则改动"
KIND_REVIEW = "需人工确认"


def _kind(cat: str) -> str:
    if cat in GROUP_FORM:
        return KIND_FORM
    if cat in GROUP_RULE:
        return KIND_RULE
    return KIND_REVIEW


@dataclass
class CellChange:
    sheet: str
    coord: str
    category: str
    old: Any
    new: Any
    value_changed: bool = True   # 有效取值是否真的变了（公式取缓存值比较）
    base: Any = None             # 原表这一格屏幕上显示的值（公式已换成算好的数）
    ratio: float | None = None   # 单价列的折扣比例
    hint: str = ""               # 整行摘要，便于定位
    reason: str = ""             # 这一处为什么改（能说清时填上）

    def brief(self, width: int = 46) -> str:
        def show(v: Any) -> str:
            if v is None or v == "":
                return "（空）"
            s = str(v).replace("\n", "⏎")
            return s if len(s) <= width else s[: width - 1] + "…"
        return f"{self.coord}｜{show(self.old)} → {show(self.new)}"

    def full(self) -> str:
        """完整原文，用于「删掉了什么」必须说清楚的场合。"""
        def show(v: Any) -> str:
            if v is None or v == "":
                return "（空）"
            s = str(v)
            body = s.replace("\n", "⏎")
            if len(body) <= 200:
                return body
            return body[:200] + f"…（共 {len(s)} 字）"
        return f"{self.coord}｜原：{show(self.old)}｜现：{show(self.new)}"


@dataclass
class RowChange:
    sheet: str
    op: str          # 删除行 / 新增行
    rows: str        # 例如 "26" 或 "18-20"
    sample: str


@dataclass
class SheetReport:
    sheet: str
    rows_equal: int = 0
    cells_equal: int = 0
    cells_compared: int = 0
    row_changes: list[RowChange] = field(default_factory=list)
    cell_changes: list[CellChange] = field(default_factory=list)
    merges_added: list[str] = field(default_factory=list)
    merges_removed: list[str] = field(default_factory=list)
    price_ratio: float | None = None

    def by_category(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for c in self.cell_changes:
            out[c.category] = out.get(c.category, 0) + 1
        return out

    def by_kind(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for c in self.cell_changes:
            k = _kind(c.category)
            out[k] = out.get(k, 0) + 1
        return out

    def review(self) -> list[CellChange]:
        return [c for c in self.cell_changes if _kind(c.category) == KIND_REVIEW]


# ── 基础工具 ────────────────────────────────────────────────────────
def _text(v: Any) -> str:
    if v is None:
        return ""
    return str(v).replace("\u3000", " ").strip()


def _norm(v: Any) -> Any:
    """比较用的归一化：None 与空串视为相同，其余原样比较（保证「一字未改」）。"""
    if v is None:
        return ""
    if isinstance(v, str):
        s = v.strip()
        return "" if s == "" else s
    return v


def _squash(v: Any) -> str:
    """把所有空白折叠成单空格，用于判断「只是排版变了」。"""
    return " ".join(_text(v).split())


def _is_formula(v: Any) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _row_signature(ws: Any, row: int, max_col: int) -> tuple:
    return tuple(_norm(ws.cell(row, c).value) for c in range(1, max_col + 1))


def _row_hint(ws: Any, row: int) -> str:
    parts = []
    for c in range(1, 9):
        t = _text(ws.cell(row, c).value)
        if t:
            parts.append(t)
    return " / ".join(parts)[:60]


def _find_header_row(ws: Any) -> int | None:
    for r in range(1, min(ws.max_row, MAX_HEADER_SCAN_ROWS) + 1):
        texts = [_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 40) + 1)]
        if any("序号" in t for t in texts) and any("产品名称" in t for t in texts):
            return r
    return None


def _locate(ws: Any, header_row: int | None) -> dict[str, int]:
    found: dict[str, int] = {}
    if header_row is None:
        return found
    for c in range(1, min(ws.max_column, 40) + 1):
        raw = _text(ws.cell(header_row, c).value)
        if not raw:
            continue
        for logical, keys in HEADER_KEYS.items():
            if logical in found:
                continue
            if any(k in raw for k in keys):
                found[logical] = c
    return found


# ── 归类 ────────────────────────────────────────────────────────────
def _base_value(old: Any, cached: Any) -> Any:
    """原表单元格的「有效取值」：公式取缓存值，普通值取自身。"""
    if _is_formula(old):
        return cached
    return old


def _classify(coord_col: int, old: Any, new: Any, cached: Any,
              cols: dict[str, int], hint: str) -> tuple[str, bool, float | None]:
    """给一处单元格改动归类，返回 (类别, 取值是否真的变了, 折扣比例)。

    分类顺序很重要：先判「值到底变没变」——
    原表公式 =+'1'!Q2 被写成同值文本 S2609-1563-14，属于形式转换，
    不能记成「备注改写」，否则真正的备注改动会被噪音淹掉。
    """
    base = _base_value(old, cached)
    # 公式没有缓存值（文件从未被 Excel 重算过）时，只能按字面判「变了」
    value_changed = _norm(base) != _norm(new) if base is not None else True

    if not value_changed:
        return CAT_FREEZE, False, None

    if cols.get("单价") == coord_col:
        ratio = None
        if isinstance(base, (int, float)) and isinstance(new, (int, float)) and base:
            ratio = round(float(new) / float(base), 4)
        return CAT_PRICE, True, ratio

    if cols.get("工分") == coord_col:
        return CAT_WORKPOINT, True, None

    # 木箱包装行只清「小计 / 工分」两列；备注文本里提到"木箱包装"不算
    if "木箱包装" in hint and coord_col in (cols.get("小计"), cols.get("工分")):
        return CAT_WOODBOX, True, None

    if cols.get("备注") == coord_col:
        return CAT_NOTE, True, None

    # 备注里的结算/折扣金额被删：原文本包含新文本（仅空白折叠后）
    if isinstance(old, str) and isinstance(new, str) and _squash(new) and _squash(new) in _squash(old):
        return CAT_NOTE_CLEAN, True, None

    # 整格就是一句报价/结算信息，被备注清理整格删掉
    if _norm(new) == "" and isinstance(old, str) and _looks_like_payment(old):
        return CAT_NOTE_CLEAN, True, None

    if _is_formula(old) and not _is_formula(new):
        return CAT_VALUE_CHANGED, True, None
    if _is_formula(old) and _is_formula(new):
        return CAT_FORMULA, True, None
    if not _is_formula(old) and _is_formula(new):
        return CAT_FORMULA, True, None
    if _norm(old) == "":
        return CAT_ADDED, True, None
    if _norm(new) == "":
        return CAT_DELETED, True, None
    return CAT_CHANGED, True, None


def _is_page_suffix_slot(old: Any, new: Any, base: Any, page: int) -> bool:
    """这格是不是被写成了页号（-工作表序号），且**取值真的变了**。

    normalize_order_numbers() 会把「订单编号」「生产单号」右侧的页号格写成 -页号。
    原值可能是旧页号、空着、或者一条 =R2 之类的公式。
    若原公式本来就算出同一个页号，那只是把公式固化成数值，不算更正。
    """
    if not isinstance(new, (int, float)) or new != page:
        return False
    if base is not None and _norm(base) == _norm(new):
        return False
    if old is None or old == "":
        return True
    if isinstance(old, (int, float)) and old < 0:
        return True
    return isinstance(old, str) and old.startswith("=")


def _explain(rep: SheetReport, name: str) -> None:
    """给「数字工作表」里能对得上规则的改动补上说明。

    normalize_order_numbers() 会把每页「订单编号 / 生产单号」右侧的页号格写成 -页号。
    绝大多数表原表就是对的，只有填错、空着或写成公式的才出现差异；
    引用该格的公式随之更新。
    """
    if not name.isdigit():
        return
    page = -int(name)
    suffix_coords: set[str] = set()
    for c in rep.cell_changes:
        if not _is_page_suffix_slot(c.old, c.new, c.base, page):
            continue
        if isinstance(c.old, (int, float)) and c.old < 0:
            c.reason = (f"订单编号后缀按工作表序号更正：原表填的是第 {-int(c.old)} 页的 "
                        f"{int(c.old)}，本页是第 {int(name)} 页，应为 {page}")
        elif c.old in (None, ""):
            c.reason = f"原表这格空着，按第 {int(name)} 页补上页号 {page}"
        else:
            shown = c.base if c.base not in (None, "") else "空"
            c.reason = f"原表这格是公式 {c.old}（显示 {shown}），改成固定的页号 {page}"
        c.category = CAT_PAGE_SUFFIX
        suffix_coords.add(c.coord)

    for c in rep.cell_changes:
        if c.reason or not isinstance(c.old, str):
            continue
        m = SINGLE_REF_RE.match(c.old.strip())
        if not m:
            continue
        target = f"{m.group(1)}{m.group(2)}"
        if target in suffix_coords:
            c.category = CAT_DERIVED
            c.reason = f"原公式 {c.old} 只引用 {target}，{target} 改后本格随之更新"


def _diff_sheet(name: str, src: Any, src_v: Any, dst: Any) -> SheetReport:
    rep = SheetReport(sheet=name)
    max_col = min(max(src.max_column, dst.max_column), 60)
    src_rows = [_row_signature(src, r, max_col) for r in range(1, src.max_row + 1)]
    dst_rows = [_row_signature(dst, r, max_col) for r in range(1, dst.max_row + 1)]

    # 列位置优先按结果表找表头；结果表找不全时，用原表表头的原表位置
    cols = _locate(dst, _find_header_row(dst))
    if not cols:
        cols = _locate(src, _find_header_row(src))

    sm = difflib.SequenceMatcher(a=src_rows, b=dst_rows, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            rep.rows_equal += i2 - i1
            rep.cells_equal += (i2 - i1) * max_col
            rep.cells_compared += (i2 - i1) * max_col
            continue

        if tag == "delete":
            span = f"{i1 + 1}" if i2 - i1 == 1 else f"{i1 + 1}-{i2}"
            rep.row_changes.append(RowChange(name, "删除行", span, _row_hint(src, i1 + 1)))
            continue

        if tag == "insert":
            span = f"{j1 + 1}" if j2 - j1 == 1 else f"{j1 + 1}-{j2}"
            rep.row_changes.append(RowChange(name, "新增行", span, _row_hint(dst, j1 + 1)))
            continue

        # replace：先按行配对做单元格级比对
        pairs = min(i2 - i1, j2 - j1)
        for k in range(pairs):
            r_src, r_dst = i1 + k + 1, j1 + k + 1
            hint = _row_hint(dst, r_dst) or _row_hint(src, r_src)
            row_cells: list[tuple[int, Any, Any, Any]] = []
            for c in range(1, max_col + 1):
                rep.cells_compared += 1
                a = src.cell(r_src, c).value
                b = dst.cell(r_dst, c).value
                if _norm(a) == _norm(b):
                    rep.cells_equal += 1
                    continue
                cached = src_v.cell(r_src, c).value if src_v is not None else None
                row_cells.append((c, a, b, cached))

            # 同一行内逐格归类
            for c, a, b, cached in row_cells:
                cat, vchg, ratio = _classify(c, a, b, cached, cols, hint)
                rep.cell_changes.append(CellChange(
                    name, f"{openpyxl.utils.get_column_letter(c)}{r_dst}",
                    cat, a, b, vchg, _base_value(a, cached), ratio, hint,
                ))

        if i2 - i1 > pairs:
            span = f"{i1 + pairs + 1}-{i2}"
            rep.row_changes.append(RowChange(name, "删除行", span, _row_hint(src, i1 + pairs + 1)))
        if j2 - j1 > pairs:
            span = f"{j1 + pairs + 1}-{j2}"
            rep.row_changes.append(RowChange(name, "新增行", span, _row_hint(dst, j1 + pairs + 1)))

    _explain(rep, name)

    s_merges = {str(m) for m in src.merged_cells.ranges}
    d_merges = {str(m) for m in dst.merged_cells.ranges}
    rep.merges_added = sorted(d_merges - s_merges)
    rep.merges_removed = sorted(s_merges - d_merges)
    return rep


# ── 汇总校验 ────────────────────────────────────────────────────────
def _finalize(reports: list[SheetReport]) -> None:
    """确认单价折扣是「统一系数」；偏离统一系数的单价改动单独挑出来复核。"""
    price_items = [c for r in reports for c in r.cell_changes if c.category == CAT_PRICE]
    ratios = [c.ratio for c in price_items if c.ratio is not None]
    if ratios:
        modal = Counter(ratios).most_common(1)[0][0]
        for r in reports:
            r.price_ratio = modal
        for c in price_items:
            if c.ratio is None or abs(c.ratio - modal) > 1e-6:
                c.category = CAT_VALUE_CHANGED


# ── 报告 ────────────────────────────────────────────────────────────
def render(reports: list[SheetReport], limit: int = 8) -> str:
    L: list[str] = []
    compared = sum(r.cells_compared for r in reports)
    equal = sum(r.cells_equal for r in reports)
    changed = sum(len(r.cell_changes) for r in reports)
    row_changes = sum(len(r.row_changes) for r in reports)
    touched = [r for r in reports if r.cell_changes or r.row_changes
               or r.merges_added or r.merges_removed]

    kinds: dict[str, int] = {}
    cats: dict[str, int] = {}
    for r in reports:
        for k, v in r.by_kind().items():
            kinds[k] = kinds.get(k, 0) + v
        for k, v in r.by_category().items():
            cats[k] = cats.get(k, 0) + v

    pct = (equal / compared * 100) if compared else 100.0
    L.append(f"比对范围：{len(reports)} 个工作表，其中 {len(touched)} 个有改动。")
    L.append(f"共比对 {compared} 个单元格：与原表字面一致 {equal} 个（{pct:.2f}%），"
             f"存在差异 {changed} 个。")
    if row_changes:
        L.append(f"整行增删：{row_changes} 处。")
    L.append("")

    L.append("差异构成：" + "、".join(f"{k} {v} 处" for k, v in
                                 sorted(kinds.items(), key=lambda x: -x[1])))
    L.append("")

    if GROUP_RULE & set(cats):
        L.append("── 规则改动（代码里写明的业务规则，逐条可查）" + "─" * 20)
        for cat in (CAT_PRICE, CAT_WORKPOINT, CAT_NOTE, CAT_NOTE_CLEAN, CAT_WOODBOX,
                    CAT_PAGE_SUFFIX, CAT_DERIVED):
            if cats.get(cat):
                extra = ""
                if cat == CAT_PRICE:
                    ratios = [r.price_ratio for r in reports if r.price_ratio]
                    if ratios:
                        extra = f"（统一系数 ×{ratios[0]}）"
                L.append(f"  · {cat}：{cats[cat]} 处{extra}")
        L.append("")

    if GROUP_FORM & set(cats):
        L.append("── 形式转换（字面不同，取值相同）" + "─" * 28)
        for cat in (CAT_FREEZE, CAT_FORMULA):
            if cats.get(cat):
                L.append(f"  · {cat}：{cats[cat]} 处")
        L.append("")

    # 需要说明的明细
    for rep in touched:
        review = rep.review()
        notable = [c for c in rep.cell_changes
                   if c.category in (CAT_NOTE, CAT_NOTE_CLEAN, CAT_WOODBOX,
                                     CAT_DELETED, CAT_ADDED, CAT_CHANGED)
                   or c.reason]
        if not notable and not rep.row_changes and not rep.merges_added and not rep.merges_removed:
            continue
        L.append(f"■ 工作表 {rep.sheet}")
        summary = "、".join(f"{k} {v}" for k, v in rep.by_category().items())
        L.append(f"   改动构成：{summary}")
        if rep.row_changes:
            for rc in rep.row_changes:
                L.append(f"   {rc.op} {rc.rows}（{rc.sample}）")
        if rep.merges_removed:
            L.append(f"   解除合并：{', '.join(rep.merges_removed[:6])}")
        if rep.merges_added:
            L.append(f"   新增合并：{', '.join(rep.merges_added[:6])}")

        shown = 0
        for c in notable:
            if shown >= limit:
                L.append(f"   … 另有 {len(notable) - shown} 处未列出")
                break
            L.append("   - " + c.full())
            if c.reason:
                L.append(f"       理由：{c.reason}")
            shown += 1

        if review:
            L.append(f"   ⚠ 需人工确认 {len(review)} 处：")
            for c in review[:limit]:
                L.append(f"     · {c.coord}｜{c.category}｜{c.old!r} → {c.new!r}")
        L.append("")

    L.append("=" * 66)
    review_total = sum(len(r.review()) for r in reports)
    unknown = cats.get(CAT_UNKNOWN, 0)
    if review_total == 0:
        L.append(f"✅ 除上列 {changed} 处已说明的改动外，其余 {equal} 个单元格"
                 f"（{pct:.2f}%）与原表一字未改。")
    else:
        L.append(f"⚠️ 有 {review_total} 处改动不属于任何已知规则，需人工确认（见上列）。")
        if unknown:
            L.append(f"   其中 {unknown} 处未能自动归类。")
    return "\n".join(L)


def audit(source: str | Path, output: str | Path) -> list[SheetReport]:
    wb_src = openpyxl.load_workbook(str(source))
    wb_src_v = openpyxl.load_workbook(str(source), data_only=True)
    wb_dst = openpyxl.load_workbook(str(output))
    reports: list[SheetReport] = []
    for name in wb_dst.sheetnames:
        if name not in wb_src.sheetnames:
            continue
        reports.append(_diff_sheet(name, wb_src[name], wb_src_v[name], wb_dst[name]))
    _finalize(reports)
    return reports


def to_payload(reports: list[SheetReport]) -> list[dict]:
    """把审计结果摊平成可 JSON 化的结构，供保存或下游（大白话说明）消费。"""
    payload = []
    for r in reports:
        payload.append({
            "sheet": r.sheet,
            "rows_equal": r.rows_equal,
            "cells_equal": r.cells_equal,
            "cells_compared": r.cells_compared,
            "price_ratio": r.price_ratio,
            "row_changes": [asdict(x) for x in r.row_changes],
            "cell_changes": [asdict(x) for x in r.cell_changes],
            "merges_added": r.merges_added,
            "merges_removed": r.merges_removed,
            "by_category": r.by_category(),
        })
    return payload


def write_json(reports: list[SheetReport], path: str | Path) -> None:
    Path(path).write_text(
        json.dumps(to_payload(reports), ensure_ascii=False, indent=2), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="下车间改动审计：逐格比对原表与结果")
    ap.add_argument("source")
    ap.add_argument("output")
    ap.add_argument("--json")
    ap.add_argument("--limit", type=int, default=8)
    args = ap.parse_args(argv)

    for p in (args.source, args.output):
        if not Path(p).exists():
            print(f"文件不存在：{p}")
            return 2

    reports = audit(args.source, args.output)
    print(render(reports, args.limit))

    if args.json:
        write_json(reports, args.json)
        print(f"\n已保存 JSON：{args.json}")
    # 有无法解释的改动就返回非 0，方便外部脚本拦截
    return 0 if not any(r.review() for r in reports) else 1


if __name__ == "__main__":
    sys.exit(main())
