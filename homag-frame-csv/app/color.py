# -*- coding: utf-8 -*-
"""PY/ZKY → YSM 对照（本服务自带一份，不引用联思/PVC 仓库）。"""

from __future__ import annotations

import re

PY_TO_YSM: dict[str, str] = {
    "PY01": "YSM262-3",
    "PY02": "YSM262-4",
    "PY03": "YSM230-1",
    "PY04": "YSM223-1",
    "PY05": "YSM232-5",
    "PY06": "YSM230-4",
    "PY07": "YSM232-4",
    "PY08": "YSM223-4",
    "PY09": "YSM01008-49",
    "PY10": "YSM232-6",
    "PY11": "YSM159-1",
    "PY12": "YSM230-3",
    "PY13": "YSM230-5",
    "PY14": "YSM223-3",
    "PY15": "YSM232-10",
    "PY16": "YSM266-3",
    "PY17": "YSM224-15",
    "PY18": "YSM02504-51",
}

ZKY_TO_YSM: dict[str, str] = {
    "ZKY5001": "YSM261-1",
    "ZKY5002": "YSM262-2",
    "ZKY5003": "YSM235-3",
    "ZKY5004": "YSM262-3",
    "ZKY5005": "YSM262-4",
    "ZKY5006": "YSM261-5",
    "ZKY5007": "YSM230-1",
    "ZKY5008": "YSM223-1",
    "ZKY5009": "YSM232-5",
    "ZKY5010": "YSM230-4",
    "ZKY5011": "YSM232-4",
    "ZKY5012": "YSM223-4",
    "ZKY5013": "YSM232-6",
    "ZKY5014": "YSM159-1",
    "ZKY5015": "YSM230-3",
    "ZKY5016": "YSM230-5",
    "ZKY5017": "YSM223-3",
    "ZKY5018": "YSM232-10",
    "ZKY5019": "YSM266-3",
    "ZKY5020": "YSM224-17",
    "ZKY5021": "YSM260-1",
    "ZKY5022": "YSM224-15",
    "ZKY5023": "YSM8897",
    "ZKY5024": "YSM01008-49",
    "ZKY5025": "YSM02504-51",
    "ZKY5026": "YSM034015",
}

COLOR_SUFFIXES = ["多层加密板", "多层加密", "黑碳晶板", "黑碳晶"]
MATERIAL_WORDS = ["密度板"] + COLOR_SUFFIXES


def _strip_batch_prefix(value: str) -> str:
    s = str(value or "").strip()
    s = re.sub(r"^(P?8[-.]\d+\w*(\s*平板)?)\s*/?", "", s)
    return s.strip()


def _extract_suffix(s: str) -> tuple[str, str]:
    suffix = ""
    for kw in MATERIAL_WORDS:
        if kw in s:
            suffix = kw
            s = s.replace(kw, "")
    return s.strip(), suffix


def _lookup_ysm(code: str) -> str | None:
    key = code.upper().replace("-", "")
    if key.startswith("PY"):
        return PY_TO_YSM.get(key)
    if key.startswith("ZKY"):
        return ZKY_TO_YSM.get(key)
    return None


def _format_ysm(raw: str, suffix: str) -> str:
    m = re.match(r"^(YSM)[-\s]*([\d-]+)(.*)$", raw, re.IGNORECASE)
    if not m:
        return f"{raw}{suffix}" if suffix else raw
    digits = m.group(2).replace(" ", "")
    rest = m.group(3).strip()
    material = suffix
    if not material:
        for kw in MATERIAL_WORDS:
            if kw in rest:
                material = kw
                break
    return f"YSM{digits}{material}"


def convert(value: str) -> tuple[str, str]:
    """返回 (对照后的 YSM, 原颜色原文)。未知编码保持原样。"""
    original = _strip_batch_prefix(value)
    if not original:
        return "", ""

    body, suffix = _extract_suffix(original)

    m = re.match(r"^(PY\d+|ZKY-?\d+)(?:-.*)?$", body, re.IGNORECASE)
    if m:
        ysm = _lookup_ysm(m.group(1))
        if ysm:
            return f"{ysm}{suffix}", original
        return f"{m.group(1).upper()}{suffix}", original

    if re.match(r"^YSM", body, re.IGNORECASE):
        return _format_ysm(body, suffix), original

    fallback = original if suffix in original else f"{original}{suffix}"
    return fallback, original


def split_batch_and_color(value: str) -> tuple[str, str]:
    """拆 H 列：'8-14G 平板/YSM232-6' → ('8-14G 平板', 'YSM232-6')。无 / 则批次为空。"""
    s = str(value or "").strip()
    if "/" in s:
        left, right = s.split("/", 1)
        return left.strip(), right.strip()
    return "", s
