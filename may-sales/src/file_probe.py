# -*- coding: utf-8 -*-
"""按**文件内容**识别销售部业绩核对三件套，不再只看文件名。

命名规则一改（2026-10 联思导出改名、群里随手改个名字）文件名识别就废，
所以这里直接打开 xlsx 看：sheet 名 → 表头关键字 → 文件名兜底。

零第三方依赖（只用 zipfile + 正则），网关和 worker 都能直接 import。
"""
from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path

# sharedStrings / 表头所在 sheet 的读取上限：够判断表头且不吃内存
_MAX_SHARED = 4 * 1024 * 1024
_MAX_SHEET = 512 * 1024
_MAX_SHEETS = 3


def probe(path: str) -> tuple[list[str], str]:
    """返回 (sheet 名列表, 文本样本)。读不了就返回 ([], '')。

    两种存储方式都要覆盖：Excel 导出的字符串多放 sharedStrings.xml，
    openpyxl/WPS 新版会写成 inlineStr 直接放在 worksheet 里。
    """
    p = Path(path)
    if not p.is_file():
        return [], ""
    try:
        with zipfile.ZipFile(p) as z:
            names = set(z.namelist())
            sheets: list[str] = []
            if "xl/workbook.xml" in names:
                wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore")
                sheets = [html.unescape(s) for s in
                          re.findall(r'<sheet\b[^>]*\bname="([^"]*)"', wbx)]

            parts: list[str] = []
            if "xl/sharedStrings.xml" in names:
                with z.open("xl/sharedStrings.xml") as fh:
                    parts.append(fh.read(_MAX_SHARED).decode("utf-8", "ignore"))
            # 表头永远在最前面几行，扫前几个 sheet 的开头就够
            sheet_xml = sorted(n for n in names
                               if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
            for n in sheet_xml[:_MAX_SHEETS]:
                with z.open(n) as fh:
                    parts.append(fh.read(_MAX_SHEET).decode("utf-8", "ignore"))
            return sheets, "".join(parts)
    except Exception:  # noqa: BLE001  老 .xls / 损坏文件 / 非 Excel
        return [], ""


# sheet 名命中（最强信号）
SHEET_MARKS = (
    ("zhcx", ("综合查询",)),
    ("liansi", ("经销商财务统计", "联思")),
    ("shejiang", ("奢匠下单表", "奢匠下单统计")),
)

# 表头关键字命中（sheet 名被改时的第二道）
HEADER_MARKS = (
    ("liansi", ("财务审核时间", "合同号"), ("订单类型", "经销商")),
    ("shejiang", ("实际销售额",), ("月份", "经销商")),
    ("zhcx", ("综合查询",), ()),
    ("template", ("销售部业绩",), ("核对", "业绩")),
)

# 文件名兜底（老逻辑）
NAME_MARKS = (
    ("zhcx", ("综合查询",)),
    ("liansi", ("联思", "财务统计")),
    ("shejiang", ("奢匠", "下单统计", "线下")),
    ("template", ("核对表", "待核对", "模板", "业绩数据表")),
)


def probe(path: str) -> tuple[list[str], str]:
    """返回 (sheet 名列表, 文本样本)。读不了就返回 ([], '')。

    两种存储方式都要覆盖：Excel 导出的字符串多放 sharedStrings.xml，
    openpyxl/WPS 新版会写成 inlineStr 直接放在 worksheet 里。
    """
    p = Path(path)
    if not p.is_file():
        return [], ""
    try:
        with zipfile.ZipFile(p) as z:
            names = set(z.namelist())
            sheets: list[str] = []
            if "xl/workbook.xml" in names:
                wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore")
                sheets = [html.unescape(s) for s in
                          re.findall(r'<sheet\b[^>]*\bname="([^"]*)"', wbx)]

            parts: list[str] = []
            if "xl/sharedStrings.xml" in names:
                with z.open("xl/sharedStrings.xml") as fh:
                    parts.append(fh.read(_MAX_SHARED).decode("utf-8", "ignore"))
            # 表头永远在最前面几行，扫前几个 sheet 的开头就够
            sheet_xml = sorted(n for n in names
                               if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
            for n in sheet_xml[:_MAX_SHEETS]:
                with z.open(n) as fh:
                    parts.append(fh.read(_MAX_SHEET).decode("utf-8", "ignore"))
            return sheets, "".join(parts)
    except Exception:  # noqa: BLE001  老 .xls / 损坏文件 / 非 Excel
        return [], ""


def detect_by_content(path: str) -> str:
    """只看内容，认不出返回 ''。"""
    sheets, texts = probe(path)
    if not sheets and not texts:
        return ""

    joined = " ".join(sheets)
    for kind, marks in SHEET_MARKS:
        if any(m in joined for m in marks):
            return kind

    for kind, must, any_of in HEADER_MARKS:
        if all(m in texts for m in must) and (not any_of or any(m in texts for m in any_of)):
            return kind
    return ""


def detect_by_name(filename: str) -> str:
    name = (filename or "").lower()
    for kind, marks in NAME_MARKS:
        if any(m.lower() in name for m in marks):
            return kind
    return "unknown"


def detect(path: str | None, filename: str = "") -> str:
    """内容优先、文件名兜底。path 为 None 时只看名字。"""
    if path:
        kind = detect_by_content(path)
        if kind:
            return kind
    return detect_by_name(filename)
