"""把审计结果翻译成大白话，输出给车间看的 Excel 清单。

输入：change_audit 生成的 JSON（逐格改动明细）
输出：改动说明_给车间.xlsx
  1) 「一页看懂」——几句话讲清这份文件动了什么、什么没动
  2) 「逐格清单」——一行一条，写清哪一页哪一格、原来什么、现在什么、为什么改

面向读者是车间员工，不出现"规则改动""形式转换""缓存值"这类词。
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── 大白话对照 ──────────────────────────────────────────────────────
# (清单里显示的类别名, 这一格是什么, 一句话说明模板, 要不要注意)
PLAIN: dict[str, tuple[str, str, str, str]] = {
    "单价折扣": (
        "单价打了 8 折",
        "单价（元）",
        "原来 {old} 元，打 8 折后是 {new} 元。",
        "常规计价，不用逐条说",
    ),
    "工分补填": (
        "补上了工分",
        "工分",
        "原来这格空着，现在补上了 {new} 工分（就是小计 ÷ 100）。",
        "排产、算工资按这个工分数走",
    ),
    "木箱包装行清空": (
        "木箱包装行删掉了",
        "小计（元）",
        "「木箱包装」这几行下车间不算钱，已删掉。",
        "不用管，下车间本来就不算钱",
    ),
    "备注清理（删掉结算/折扣金额）": (
        "备注里删掉了报价金额",
        "备注",
        "备注里那句讲优惠、讲折后多少钱的话被删掉了，只留车间用得上的话。",
        "看看删掉的内容车间确实用不上",
    ),
    "订单编号后缀更正": (
        "页号数字写错了，已改对",
        "订单编号后面的页号",
        "原表这格填的是 {old}，是别的页留下来的旧数字，这页应该是 {new}，已改对。",
        "只是页号数字，不影响金额",
    ),
    "连带更新（跟随被引用单元格）": (
        "跟着上面那格一起变",
        "公式格",
        "这格跟上面那格是一伙的，上面改成 {new}，它也跟着变。",
        "不用单独处理",
    ),
    "公式固化为数值（结果不变）": (
        "公式换成固定的数，数没变",
        "公式格",
        "原来这格是用公式从别的表取数，现在直接写成了数字，显示的数一模一样。",
        "不用管。以后换电脑、断网也不会变空白",
    ),
    "公式结果变了": (
        "公式算出来的数变了",
        "公式格",
        "这格原来是公式，现在固化成 {new}，跟原来的结果不一样。",
        "★ 请核对这一格对不对",
    ),
    "内容改写": ("内容被改了", "内容", "原来 {old}，现在 {new}。", "★ 请确认是不是该改"),
    "新增内容": ("原来空格，填了新内容", "内容", "原来空着，现在填了 {new}。", "请确认"),
    "删除内容": ("原来的内容被删了", "内容", "原来 {old}，现在空了。", "★ 请确认是不是该删"),
    "未归类": ("说不清为什么改", "未知", "原来 {old}，现在 {new}。", "★ 必须人工核实"),
}

ORDER = [
    "订单编号后缀更正",
    "连带更新（跟随被引用单元格）",
    "木箱包装行清空",
    "备注清理（删掉结算/折扣金额）",
    "工分补填",
    "单价折扣",
    "公式固化为数值（结果不变）",
    "内容改写",
    "新增内容",
    "删除内容",
    "公式结果变了",
    "未归类",
]

# 常规动作：属于我们自己的规矩或日常维护，不跟车间逐条解释
# （单价 8 折一户就两百多条，页号校正这种数字细节车间也用不上，
#  混进去会把真正要交代的事淹掉）
ROUTINE = {"单价折扣", "订单编号后缀更正", "连带更新（跟随被引用单元格）"}

# 需要人工确认的类别（与 change_audit.GROUP_REVIEW 对应）
GROUP_REVIEW = {"公式结果变了", "新增内容", "删除内容", "内容改写", "未归类"}
CAT_NOTE_CLEAN = "备注清理（删掉结算/折扣金额）"
CAT_WOODBOX = "木箱包装行清空"

THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=11)
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
TITLE_FONT = Font(bold=True, size=16, color="1F4E79")
H2_FONT = Font(bold=True, size=12, color="1F4E79")


def _num(v: Any) -> str:
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _show(v: Any, limit: int = 300) -> str:
    if v is None or v == "":
        return "（空）"
    s = _num(v) if isinstance(v, (int, float)) else str(v)
    s = s.replace("\r", "").replace("\n", " ⏎ ")
    s = re.sub(r"[ \u3000]{3,}", " ", s)
    return s if len(s) <= limit else s[:limit] + f"……（后面还有 {len(s) - limit} 字）"


def _row_of(coord: str) -> str:
    m = re.match(r"[A-Z]+(\d+)", coord)
    return m.group(1) if m else ""


def _squash(v: Any) -> str:
    """把空白折叠成单空格，用来判断"只是排版变了"。"""
    if v is None:
        return ""
    return " ".join(str(v).replace("\u3000", " ").split())


def _intact_text(same: int, total: int) -> str:
    """「一个字没改」占多少。

    有改动时绝不能显示 100%：14 处改动配 100% 是自相矛盾的，
    罪魁是四舍五入——99.98 会被 %.1f 说成 100.0。这里留两位小数并封顶 99.99%。
    """
    if not total or same >= total:
        return "100%"
    return f"{min(same / total * 100, 99.99):.2f}%"


def _removed_sentences(old: Any, new: Any) -> str:
    """从长备注里挑出「被删掉的那几句」，让员工一眼看到删了什么。"""
    if not isinstance(old, str) or not isinstance(new, str):
        return ""
    keep = _squash(new)
    parts = re.split(r"[。！？!?\n]+", old)
    gone = [p.strip(" ，,、;；") for p in parts if p.strip() and _squash(p) not in keep]
    return "；".join(g for g in gone if g)


def _where(it: dict) -> str:
    """「2页T9」这种位置写法，摘要里统一用。"""
    sheet = it.get("sheet")
    coord = it.get("coord") or ""
    return f"{sheet}页{coord}" if str(sheet).isdigit() else f"{sheet}{coord}"


def _note_deleted(it: dict) -> str:
    """备注这一格被删掉的原文：整格清空就取原句，只删几句就取那几句。"""
    old, new = it.get("old"), it.get("new")
    if new in (None, ""):
        return str(old) if old not in (None, "") else ""
    gone = _removed_sentences(old, new)
    if gone:
        return gone
    # 逐字改写的，两头都摆出来，别让人猜
    return f"{_show(old, 60)} → {_show(new, 60)}"


def _note_lines(items: list[dict]) -> list[str]:
    """删备注必须把原文写全。

    备注里那几句话是讲优惠、讲折后多少钱的，车间要照着核对，
    只报「删备注 1 处」等于什么都没说。这一行不省、不截成一句摘要。
    """
    groups: dict[str, list[str]] = {}
    for it in items:
        if str(it.get("category") or "") != CAT_NOTE_CLEAN:
            continue
        text = _note_deleted(it)
        if not text or not text.strip():
            continue
        groups.setdefault(text.strip(), []).append(_where(it))

    lines = []
    for text, places in groups.items():
        spot = "、".join(dict.fromkeys(places))
        lines.append(f"🗒 删备注（{spot}）：「{_show(text, 200)}」")
    if len(lines) > MAX_NOTE_LINES:
        extra = len(lines) - MAX_NOTE_LINES
        lines = lines[:MAX_NOTE_LINES] + [f"🗒 另有 {extra} 处备注，原文见改动清单"]
    return lines


def _plain(cat: str, old: Any, new: Any, base: Any) -> tuple[str, str, str, str, Any, Any]:
    """把技术类别翻成大白话。

    展示「原来」时用 base（屏幕上看到的数），不用原始公式——
    车间看的是表格画面，不是公式栏，写 =398+60+90+40 只会让人看不懂。
    """
    name, what, tpl, note = PLAIN.get(cat, ("其它改动", "未知", "原来 {old}，现在 {new}。", "请确认"))
    shown_old = base if base not in (None, "") else old
    shown_new = new
    text = tpl.format(old=_show(shown_old, 80), new=_show(new, 80))

    if cat == "工分补填" and isinstance(new, str) and new.startswith("="):
        # 工分是新写进去的公式，文件里还没有算好的数；直接说清算法，别甩公式
        shown_new = "自动算（小计 ÷ 100）"
        text = "原来这格空着，现在按「小计 ÷ 100」自动补上了工分。"

    if cat == "备注清理（删掉结算/折扣金额）":
        gone = _removed_sentences(old, new)
        if gone:
            text = f"备注里删掉了这一句：「{gone}」"

    if cat == "公式固化为数值（结果不变）" and shown_old not in (None, ""):
        # 原来和现在数字一样，员工会疑惑"一样还叫改？"，直接把"原来是公式"标出来
        shown_old = f"{_show(shown_old, 80)}（原来是公式算的）"

    return name, what, text, note, shown_old, shown_new


# 发到群里的一句话说明：每类只讲"为什么改、要不要注意"，不逐格列
BRIEF: dict[str, tuple[str, str]] = {
    "工分补填": ("空着的工分补上了", "按「小计 ÷ 100」自动补的。排产、算工资用这个数。"),
    "木箱包装行清空": ("木箱包装行删了", "这几行下车间不算钱，已经删掉。"),
    "备注清理（删掉结算/折扣金额）": ("备注里删了话", "讲优惠、讲折后多少钱的那些话删掉了，其余原样留着。"),
    "公式固化为数值（结果不变）": ("公式换成固定数字", "原来这些格子是从别的表取数，现在直接写死成数字。屏幕上看到的数一模一样，好处是换电脑、断网都不会变空白。"),
    "内容改写": ("有格子内容变了", "★ 和原表不一样，需要确认是不是该改。"),
    "新增内容": ("补了原来空着的内容", "★ 需要确认。"),
    "删除内容": ("删掉了原来的内容", "★ 需要确认是不是该删。"),
    "公式结果变了": ("公式算出来的数变了", "★ 需要核对。"),
    "未归类": ("有几处说不清为什么改", "★ 必须人工核实。"),
}


def order_hint(name: str) -> str:
    """从文件名里取单号，用作说明的标题。"""
    m = re.search(r"[A-Za-z]\d{4}(?:-\d+)+", name)
    return f"{m.group(0)} " if m else ""


# 群里一行摘要用的极短名（build_brief）。车间扫一眼就知道动了哪儿，
# 不用读整段话——要逐格抠细节的，去看改动清单。
ABBR: dict[str, str] = {
    "单价折扣": "单价8折",
    "工分补填": "补工分",
    "木箱包装行清空": "删木箱行",
    "备注清理（删掉结算/折扣金额）": "删备注",
    "订单编号后缀更正": "页号校正",
    "连带更新（跟随被引用单元格）": "公式跟随",
    "公式固化为数值（结果不变）": "公式定值",
    "内容改写": "内容变",
    "新增内容": "新增",
    "删除内容": "删除",
    "公式结果变了": "数值变",
    "未归类": "待核",
}

# 整份文件统一的机械动作，带数字只会更吵
# （单价 8 折一户两百多处、公式定值上百格，说"多少处"对车间没用）
NO_COUNT = {
    "单价折扣",
    "连带更新（跟随被引用单元格）",
    "公式固化为数值（结果不变）",
    "订单编号后缀更正",
}
# 纯内部连带，车间根本不用知道
SKIP_BRIEF = {"连带更新（跟随被引用单元格）"}
MAX_BRIEF_ITEMS = 4
# 删备注的原文单独成行，最多这么几行（正常一单 1-2 处，不会触发）
MAX_NOTE_LINES = 8

# 摘要里的先后：先说要核对的、真删改的，全局规则（8 折、页号）压最后。
# 顺手扫一眼就先看到"哪儿动了"，而不是"单价又打折了"。
BRIEF_ORDER = [
    "未归类",
    "公式结果变了",
    "内容改写",
    "新增内容",
    "删除内容",
    "木箱包装行清空",
    "备注清理（删掉结算/折扣金额）",
    "工分补填",
    "公式固化为数值（结果不变）",
    "订单编号后缀更正",
    "单价折扣",
]


def build_brief(items: list[dict], stats: dict[str, Any] | None = None,
                hint: str = "") -> tuple[str, str]:
    """把改动压成两行：①「单号 · 改了什么」②「要核对的几格」。

    群消息只发这两行。stats 不参与计算，保留参数只为和 build_text 调用姿势一致。
    """
    counts: dict[str, int] = {}
    for it in items:
        cat = str(it.get("category") or "")
        if not cat or cat in SKIP_BRIEF:
            continue
        counts[cat] = counts.get(cat, 0) + 1

    parts: list[str] = []
    for cat in BRIEF_ORDER:
        n = counts.get(cat)
        if not n:
            continue
        name = ABBR.get(cat, cat)
        parts.append(name if cat in NO_COUNT else f"{name}{n}处")
    for cat, n in counts.items():          # 万一以后加了新类别，别吞掉
        if cat not in BRIEF_ORDER:
            parts.append(f"{ABBR.get(cat, cat)}{n}处")

    if len(parts) > MAX_BRIEF_ITEMS:
        extra = len(parts) - MAX_BRIEF_ITEMS
        parts = parts[:MAX_BRIEF_ITEMS] + [f"等{extra}类"]
    body = " · ".join(parts) if parts else "没改动"
    title = (hint or "").strip()
    line = f"{title} · {body}" if title else body

    # 删备注的原文另起一行，跟摘要行一起发（网关只在首行加 ✅，其余原样贴出）
    notes = _note_lines(items)
    if notes:
        line = "\n".join([line] + notes)

    review = [it for it in items if str(it.get("category") or "") in GROUP_REVIEW]
    if not review:
        return line, ""
    spots = [_where(c) for c in review[:4]]
    tail = f" 等{len(review)}处" if len(review) > len(spots) else ""
    return line, f"⚠ 核对{len(review)}处：{'、'.join(spots)}{tail}"


def _deleted_line(it: dict) -> str:
    """一行说清「删掉了哪一句」，让收的人不用去翻原表。"""
    sheet = it["sheet"]
    where = f"第 {sheet} 页 {it['coord']}" if str(sheet).isdigit() else f"{sheet} 的 {it['coord']}"
    old, new = it.get("old"), it.get("new")
    # 原值是公式时用它的显示值——车间看的是屏幕上那个数，不是公式
    shown_old = it.get("base") if it.get("base") not in (None, "") else old

    if new in (None, ""):                      # 整格被清空
        if isinstance(shown_old, (int, float)):
            return f"{where}：原来 {_num(shown_old)}，现在清空了"
        if isinstance(shown_old, str) and shown_old.strip():
            return f"{where}：原来整格写的「{_show(shown_old, 70)}」，现在清空了"
        return f"{where}：原来就是空的，现在清空了"

    gone = _removed_sentences(old, new)        # 只删掉其中几句
    if gone:
        return f"{where}：删掉了「{_show(gone, 70)}」"
    return f"{where}：原来 {_show(shown_old, 30)}，现在 {_show(new, 30)}"


def build_text(items: list[dict], stats: dict[str, Any], hint: str = "") -> str:
    """产出可直接发群的一段话。按类别讲，不逐格列，控制在手机一屏能看完。

    「需要人工确认」的条目必须写清是哪一页哪一格、原来什么、现在什么——
    只写"★ 请核对"，收的人根本不知道要去核对什么。
    """
    listed = [it for it in items if it["category"] not in ROUTINE]
    review = [it for it in listed if it["category"] in GROUP_REVIEW]
    normal = [it for it in listed if it["category"] not in GROUP_REVIEW]

    total_cells = stats["compared"]
    same = stats["equal"]

    title = f"【{hint}下车间文件 · 改动说明】" if hint else "【下车间文件 · 改动说明】"
    head = f"这次跟报价单比，有 {len(listed)} 处要跟大家说清楚，其余一个字没改。"

    counts: dict[str, int] = {}
    for it in normal:
        counts[it["category"]] = counts.get(it["category"], 0) + 1

    lines = [title, "", head, ""]
    idx = 0
    for cat in ORDER:
        if not counts.get(cat):
            continue
        idx += 1

        # 删了话的，必须把删掉的原文列出来——只说"删掉了报价那句"等于没说
        if cat == CAT_NOTE_CLEAN or cat == CAT_WOODBOX:
            hits = [it for it in normal if it["category"] == cat]
            if cat == CAT_WOODBOX:
                lines.append(f"{idx}）木箱包装行删了（{len(hits)} 处）")
                lines.append("这几行下车间不算钱，已经清掉：")
            else:
                lines.append(f"{idx}）备注里删了话（{len(hits)} 处）")
            for it in hits[:6]:
                lines.append("   · " + _deleted_line(it))
            if len(hits) > 6:
                lines.append(f"   · 另外还有 {len(hits) - 6} 处，见清单")
            lines.append("")
            continue

        name, brief = BRIEF.get(cat, (PLAIN.get(cat, (cat,))[0], "见清单。"))
        lines.append(f"{idx}）{name}（{counts[cat]} 处）")
        lines.append(brief)
        lines.append("")

    if review:
        idx += 1
        lines.append(f"{idx}）有 {len(review)} 处要人工确认")
        for c in review[:8]:
            sheet = c["sheet"]
            where = f"第 {sheet} 页 {c['coord']}" if str(sheet).isdigit() else f"{sheet} 的 {c['coord']}"
            old = c.get("base") if c.get("base") not in (None, "") else c.get("old")
            old_txt = _show(old, 34)
            raw = c.get("old")
            if isinstance(raw, str) and raw.startswith("="):
                # 原来是个公式：既给显示值也给公式，方便照着原表找
                old_txt = f"{old_txt}（原公式 {raw}）"
            lines.append(f"   · {where}：原来 {old_txt}，现在 {_show(c['new'], 34)}")
        if len(review) > 8:
            lines.append(f"   · 另外还有 {len(review) - 8} 处，见清单")
        lines.append("   这几处不属于常规规则，麻烦对一下再发车间。")
        lines.append("")

    if not idx:
        lines.append("这次没有需要跟大家交代的改动。")
        lines.append("")

    lines.append(f"其余没动：剩下 {same} 个格子（{_intact_text(same, total_cells)}）一个字没改。")
    lines.append("")
    if idx:
        lines.append("哪一格看不明白，报「第几页、哪个格子」就行，比如「第 1 页 L23」。")
    return "\n".join(lines)


def build(items: list[dict], stats: dict[str, Any]) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    routine_n = sum(1 for it in items if it["category"] in ROUTINE)
    listed_n = len(items) - routine_n

    # ── 第一页：一页看懂 ──────────────────────────────────────────
    ws = wb.active
    ws.title = "一页看懂"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 62

    ws["B2"] = "这份下车间文件，我们改了什么"
    ws["B2"].font = TITLE_FONT
    ws.row_dimensions[2].height = 28

    total_cells = stats["compared"]
    same = stats["equal"]
    diff = len(items)

    intro = [
        ("一共看了", f"{stats['sheets']} 张表、{total_cells} 个格子"),
        ("没动的", f"{same} 个格子（{_intact_text(same, total_cells)}）一个字都没改"),
        ("常规动作", f"{routine_n} 个格子（单价 8 折、页号校正）按老规矩办，不单列"),
        ("要交代的", f"{listed_n} 个格子，全部列在「逐格清单」里，一行一条"),
        ("合并格 / 行列位置", "没动过，跟原表一模一样"),
    ]
    r = 4
    for k, v in intro:
        ws.cell(r, 2, k).font = Font(bold=True, size=11)
        ws.cell(r, 3, v).alignment = Alignment(horizontal="left")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 20
        r += 1

    r += 1
    ws.cell(r, 2, "按类别看").font = H2_FONT
    r += 1
    for i, h in enumerate(["改了什么", "几处", "给大家的说法"], start=2):
        c = ws.cell(r, i, h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
    r += 1

    counts: dict[str, int] = {}
    for it in items:
        counts[it["category"]] = counts.get(it["category"], 0) + 1
    for cat in ORDER:
        if cat not in counts or cat in ROUTINE:
            continue
        name, what, _, note = PLAIN.get(cat, ("其它", "未知", "", ""))
        ws.cell(r, 2, name).border = BORDER
        ws.cell(r, 3, counts[cat]).border = BORDER
        ws.cell(r, 3).alignment = Alignment(horizontal="center")
        c = ws.cell(r, 4, note)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if note.startswith("★"):
            for col in range(2, 5):
                ws.cell(r, col).fill = WARN_FILL
        ws.row_dimensions[r].height = 22
        r += 1

    ws.cell(r, 2, "合计").font = Font(bold=True)
    ws.cell(r, 2).border = BORDER
    ws.cell(r, 3, listed_n).font = Font(bold=True)
    ws.cell(r, 3).border = BORDER
    ws.cell(r, 3).alignment = Alignment(horizontal="center")
    ws.cell(r, 4).border = BORDER

    r += 2
    ws.cell(r, 2, "打个比方").font = H2_FONT
    r += 1
    notes = [
        "· 有几个空着的工分格，我们按「小计 ÷ 100」补上了，排产用这个数。",
        "· 「木箱包装」那几行下车间不算钱，已经删掉。",
        "· 备注里讲优惠、讲折后多少钱的那句话删掉了，车间用不上。",
        "· 有些格子原来是用公式从别的表取数，现在直接写成了数字，看着一模一样，好处是换了电脑也不会变空白。",
    ]
    for n in notes:
        ws.cell(r, 2, n)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1

    r += 1
    ws.cell(r, 2, "对不上就找我们").font = H2_FONT
    r += 1
    ws.cell(r, 2, "清单里每一行都写了改了什么、为什么改。哪一格看不明白，"
                  "把「页」和「格子」报给我们（例如：第 1 页 L23），我们对着这行解释。")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 34

    # ── 第二页：逐格清单 ──────────────────────────────────────────
    ws2 = wb.create_sheet("逐格清单")
    headers = ["序号", "页（工作表）", "格子", "这一格是什么", "原来写的", "现在写的",
               "一句话说明", "要注意吗"]
    widths = [6, 14, 8, 16, 34, 34, 46, 34]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws2.cell(1, i, h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.row_dimensions[1].height = 26
    ws2.freeze_panes = "A2"
    listed = [it for it in items if it["category"] not in ROUTINE]
    ws2.auto_filter.ref = f"A1:H{len(listed) + 1}"

    rows = []
    for it in listed:
        name, what, text, note, shown_old, shown_new = _plain(
            it["category"], it["old"], it["new"], it.get("base"))
        rows.append({
            "sheet": it["sheet"], "coord": it["coord"], "what": what,
            "old": shown_old, "new": shown_new, "text": text,
            "note": note, "cat": it["category"],
        })

    # 按改动类型分组（同一类排在一起，员工看起来是一条一条的整块），组内按页码、行号
    cat_rank = {c: i for i, c in enumerate(ORDER)}
    rows.sort(key=lambda x: (
        cat_rank.get(x["cat"], 99),
        (0, int(x["sheet"])) if x["sheet"].isdigit() else (1, x["sheet"]),
        int(_row_of(x["coord"]) or 0),
    ))

    r = 2
    for idx, it in enumerate(rows, start=1):
        vals = [idx, f"第 {it['sheet']} 页" if it["sheet"].isdigit() else it["sheet"],
                it["coord"], it["what"], _show(it["old"]), _show(it["new"]),
                it["text"], it["note"]]
        for col, v in enumerate(vals, start=1):
            c = ws2.cell(r, col, v)
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top",
                                    horizontal="center" if col in (1, 3) else "left")
            if col >= 7:
                c.font = Font(size=10)
        if it["note"].startswith("★"):
            for col in range(1, 9):
                ws2.cell(r, col).fill = WARN_FILL
        # 按字数粗估行高，避免长备注被压成一行
        longest = max(len(str(v)) for v in vals[3:])
        ws2.row_dimensions[r].height = min(150, max(20, 15 * (longest // 40 + 1)))
        r += 1

    # ── 第三页：按页汇总（每页要交代几处，方便一页一页核） ────────
    ws3 = wb.create_sheet("按页汇总")

    by_sheet: dict[str, dict[str, int]] = {}
    warn_by_sheet: dict[str, int] = {}
    for it in listed:
        d = by_sheet.setdefault(it["sheet"], {})
        d[it["category"]] = d.get(it["category"], 0) + 1
        if PLAIN.get(it["category"], ("", "", "", ""))[3].startswith("★"):
            warn_by_sheet[it["sheet"]] = warn_by_sheet.get(it["sheet"], 0) + 1

    has_warn = bool(warn_by_sheet)
    heads = ["页（工作表）", "要交代几处", "都改了什么"] + (["其中要看的"] if has_warn else [])
    hw = [18, 12, 62] + ([12] if has_warn else [])
    for i, (h, w) in enumerate(zip(heads, hw), start=1):
        c = ws3.cell(1, i, h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A2"

    r = 2
    for sheet, d in by_sheet.items():
        parts = [f"{PLAIN.get(c, (c,))[0]} {n} 处"
                 for c, n in sorted(d.items(), key=lambda x: ORDER.index(x[0]) if x[0] in ORDER else 99)]
        vals = [f"第 {sheet} 页" if sheet.isdigit() else sheet, sum(d.values()),
                "，".join(parts)]
        if has_warn:
            vals.append(warn_by_sheet.get(sheet) or "")
        for col, v in enumerate(vals, start=1):
            c = ws3.cell(r, col, v)
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                    horizontal="center" if col in (2, 4) else "left")
        ws3.row_dimensions[r].height = max(20, 15 * (len(vals[2]) // 38 + 1))
        r += 1

    return wb


def summarize(items: list[dict], reports: list[dict]) -> dict[str, Any]:
    return {
        "sheets": len(reports),
        "compared": sum(r["cells_compared"] for r in reports),
        "equal": sum(r["cells_equal"] for r in reports),
    }


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="生成给车间看的大白话改动说明（Excel 清单 或 直接发群的一段话）")
    ap.add_argument("audit_json")
    ap.add_argument("-o", "--output", default="改动说明_给车间.xlsx")
    ap.add_argument("--text", action="store_true", help="只输出可以直接发群的一段话，不做 Excel")
    ap.add_argument("--hint", default="", help="标题前的单号，例如 S2609-1563 ")
    args = ap.parse_args(argv)

    reports = json.loads(Path(args.audit_json).read_text(encoding="utf-8"))
    items = [c for r in reports for c in r["cell_changes"]]
    stats = summarize(items, reports)

    if args.text:
        print(build_text(items, stats, args.hint))
        return 0

    wb = build(items, stats)
    wb.save(args.output)
    listed_n = sum(1 for it in items if it["category"] not in ROUTINE)
    print(f"已生成：{args.output}")
    print(f"  工作表 {len(reports)} 张，逐格清单 {listed_n} 条（另有 "
          f"{len(items) - listed_n} 条常规动作未列入）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
