#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""红字行豁免 + 剔除行记录。

两件事：
1) 豁免：PVC 分类流程把"标红"的行（字体 colour_index == 10）当成"不做"直接剔除，
   但"封边条"这类特殊工件（如 5 米封门扇用的封边条）必须保留进分类输出。
2) 记录：被剔除的行（以及被豁免保留的行）都要能回传到飞书卡片上，让人看得见。

事件收集用 threading.local，处理进程内按线程隔离，process_input 开始时 reset。
"""
import re
import threading

# 命中这些关键字的行，即使标红也不剔除
KEEP_ROW_KEYWORDS = ('封边条',)

_local = threading.local()

_BARCODE_RE = re.compile(r'^[A-Za-z]{1,3}\d{6,}$')
_ORDER_RE = re.compile(r'^[A-Za-z]{0,4}[\d\-]{5,}$')


def is_keep_row(values):
    """values = 该行所有单元格的值；命中任一豁免关键字返回 True。"""
    for v in values:
        if v is None:
            continue
        s = str(v)
        for kw in KEEP_ROW_KEYWORDS:
            if kw in s:
                return True
    return False


def read_row_values(sheet, row_idx):
    """取出整行原始值（用于豁免判断）。"""
    try:
        return [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]
    except Exception:
        return []


# ==================== 事件记录 ====================

def reset_events():
    """每次处理一个文件前调用，清空本线程记录。"""
    _local.excluded = []
    _local.kept = []


def _bucket(name):
    lst = getattr(_local, name, None)
    if lst is None:
        lst = []
        setattr(_local, name, lst)
    return lst


def row_label(values):
    """从整行值里挑出便于人看的字段：工件名称 / 条形码 / 生产号。"""
    name = ''
    # 管线中间/输出文件第 22 列(index 21)是「原工件名称」整串，优先用它
    if len(values) > 21 and values[21] is not None and str(values[21]).strip():
        name = str(values[21]).strip()
    elif values:
        name = str(values[0]).strip() if values[0] is not None else ''
        if len(values) > 1 and values[1] is not None and str(values[1]).strip() and '_' not in name:
            name = f"{name}_{str(values[1]).strip()}"
    barcode = ''
    for v in values:
        s = str(v).strip() if v is not None else ''
        if s and not barcode and _BARCODE_RE.match(s):
            barcode = s
    order = ''
    if len(values) > 8 and values[8] is not None:
        s = str(values[8]).strip()
        if len(s) >= 5 and re.search(r'\d', s):
            order = s
    return name, order, barcode


def _event(sheet, values, reason):
    name, order, barcode = row_label(values)
    return {'sheet': sheet or '', 'name': name, 'order': order,
            'barcode': barcode, 'reason': reason}


def add_excluded(sheet, values, reason='红字标记'):
    """记录一行被剔除。"""
    _bucket('excluded').append(_event(sheet, values, reason))


def add_kept(sheet, values, reason='标红但含封边条'):
    """记录一行标红但仍被保留。"""
    _bucket('kept').append(_event(sheet, values, reason))


def _fmt_event(ev):
    parts = []
    if ev['name']:
        parts.append(ev['name'])
    if ev['barcode']:
        parts.append('条码' + ev['barcode'])
    if ev['order']:
        parts.append('单号' + ev['order'])
    return '/'.join(parts) if parts else '(无名称)'


def _bullets(events, max_items):
    items = [_fmt_event(ev) for ev in events[:max_items]]
    text = '；'.join(items)
    if len(events) > max_items:
        text += f'；等共 {len(events)} 行'
    return text


def get_events():
    return {'excluded': list(_bucket('excluded')), 'kept': list(_bucket('kept'))}


def format_events(max_items=12):
    """生成回传卡片上要显示的文字行（list[str]）。"""
    excluded = _bucket('excluded')
    kept = _bucket('kept')
    lines = []
    if kept:
        lines.append(f'✅ 标红但含「封边条」已保留 {len(kept)} 行：{_bullets(kept, max_items)}')
    if excluded:
        reasons = {}
        for ev in excluded:
            reasons[ev['reason']] = reasons.get(ev['reason'], 0) + 1
        reason_text = '、'.join(f'{k}×{v}' for k, v in reasons.items())
        lines.append(f'🗑 已剔除 {len(excluded)} 行（{reason_text}）：{_bullets(excluded, max_items)}')
    return lines
