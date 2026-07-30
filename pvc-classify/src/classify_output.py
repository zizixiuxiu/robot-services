#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分类输出脚本：读取源文件（门框/哑口套/护墙），按颜色/厚度/工艺分类，
输出到以颜色命名的 .xls 文件中。
"""

import os
import sys
import shutil
import json
import re
import struct
import xlrd
import xlwt
from xlutils.copy import copy

OVERSIZE_LIMIT = 2420


def get_date_code_from_filename(filename):
    base = os.path.basename(filename)
    # Pattern 1: W 在日期前或后
    #   W6-14_for_classify.xls -> W6-14
    #   6-14W平板.xls / P6-14W平板.xls -> W6-14
    m = re.search(r'W(\d+-\d+)', base, re.IGNORECASE)
    if m:
        return 'W' + m.group(1)
    m = re.search(r'(\d+-\d+)[^\d-]*W', base, re.IGNORECASE)
    if m:
        return 'W' + m.group(1)
    # Pattern 2: P5-24.xls / P6-10_xxx.xls -> 5-24 / 6-10
    m = re.match(r'^P(\d+-\d+)', base)
    if m:
        return m.group(1)
    # Pattern 3: 5-24_for_classify.xls -> 5-24
    m = re.search(r'(\d+-\d+)', base)
    if m:
        return m.group(1)
    return None

def read_sheet_data(xls_path, sheet_name):
    """Read data rows from specified sheet, skipping header and red-font rows."""
    book = xlrd.open_workbook(xls_path, formatting_info=True)
    sheet = book.sheet_by_name(sheet_name)
    rows = []
    skipped = []
    for i in range(1, sheet.nrows):
        row = [sheet.cell_value(i, j) for j in range(sheet.ncols)]
        if len(row) > 9 and row[9]:
            # Check if row is marked red (font colour_index == 10)
            is_red = False
            for j in range(min(len(row), sheet.ncols)):
                xf = sheet.cell_xf_index(i, j)
                fmt = book.xf_list[xf]
                font = book.font_list[fmt.font_index]
                if font.colour_index == 10:
                    is_red = True
                    break
            if is_red:
                skipped.append(str(row[9]))
                continue
            rows.append(row)
    if skipped:
        print(f"  Auto-skipped {len(skipped)} red-marked rows: {', '.join(sorted(set(skipped)))}")
    return rows

def load_skip_list(skip_file):
    """Load skip order IDs from text file (one per line)."""
    skips = set()
    if skip_file and os.path.exists(skip_file):
        with open(skip_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    skips.add(line)
    return skips

def filter_skipped_rows(data_rows, skip_set):
    """Remove rows whose order ID is in skip_set."""
    if not skip_set:
        return data_rows, []
    filtered = []
    skipped_ids = []
    for row in data_rows:
        order_id = row[9] if len(row) > 9 else ''
        if order_id in skip_set:
            skipped_ids.append(order_id)
        else:
            filtered.append(row)
    return filtered, skipped_ids

def extract_base_color(color):
    """Extract base color name for filename."""
    if not color:
        return ''
    color = str(color).strip()
    color = re.sub(r'[（(]\s*(?:多层加密)?\s*[）)]', '', color)
    # Remove common suffixes that are descriptive but not color identifiers
    for suffix in ['门套门扇']:
        if color.endswith(suffix):
            color = color[:-len(suffix)]
    # 多层加密 is a material feature, not part of color name
    if '多层加密' in color:
        color = color.replace('多层加密', '')
    m = re.match(r'^([A-Z]+\d+)-([\u4e00-\u9fff]+)$', color)
    if m:
        return m.group(1)
    m = re.match(r'^(ZKY)-(\d+)$', color)
    if m:
        return m.group(1) + m.group(2)
    return color

def normalize_color_for_lookup(color):
    """Normalize color for family lookup."""
    if not color:
        return ''
    c = re.sub(r'[（(]\s*(?:多层加密)?\s*[）)]', '', str(color).strip()).upper()
    c = c.replace('多层加密', '')
    m = re.match(r'^YSM-?(\d+)-(\d+)$', c)
    if m:
        return f'YSM-{m.group(1)}-{m.group(2)}'
    m = re.match(r'^YSM-(\d+)$', c)
    if m:
        return f'YSM{m.group(1)}'
    m = re.match(r'^YSM(\d+)$', c)
    if m:
        return f'YSM{m.group(1)}'
    m = re.match(r'^ZKY-(\d+)$', c)
    if m:
        return f'ZKY{m.group(1)}'
    m = re.match(r'^([A-Z]+\d+)$', c)
    if m:
        return m.group(1)
    return c

def load_color_families(json_path=None):
    """Load color family mapping from JSON (new format with YSM defaults)."""
    if json_path and os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        families = {}
        defaults = {}
        for k, v in data.items():
            norm_k = normalize_color_for_lookup(k)
            if isinstance(v, dict) and 'members' in v:
                members = set(v['members'])
                families[norm_k] = members
                # Use the default name from JSON for all color families
                defaults[norm_k] = v.get('default', k)
            else:
                # Old format compatibility
                members = set(v) if isinstance(v, (list, tuple)) else {v}
                families[norm_k] = members
                defaults[norm_k] = k
        return families, defaults
    return {}, {}


def build_member_to_family(color_families):
    """Map every normalized color alias to its normalized family key."""
    member_to_family = {}
    for family_key, members in color_families.items():
        norm_family = normalize_color_for_lookup(family_key)
        if norm_family not in member_to_family:
            member_to_family[norm_family] = family_key
        for member in members:
            norm_member = normalize_color_for_lookup(member)
            if norm_member not in member_to_family:
                member_to_family[norm_member] = family_key
    return member_to_family


def sort_key_for_row(row):
    """Sort key: normal production first, then by order ID."""
    production_type = row[14] if len(row) > 14 else ''
    order_id = row[9] if len(row) > 9 else ''
    is_urgent = 1 if '加急' in str(production_type) else 0
    return (is_urgent, str(order_id))


def parse_thickness(value, default):
    try:
        return int(float(value)) if value not in (None, '') else default
    except (ValueError, TypeError):
        return default


def non_default_thickness_suffix(value, default):
    thickness = parse_thickness(value, default)
    return '' if thickness == default else f'{thickness}厚'


def craft_from_color(color):
    return '多层加密' if '多层加密' in str(color) else ''


def is_heitanjing(row):
    """黑碳晶可能写在颜色列或材质列。"""
    color = str(row[7]) if len(row) > 7 else ''
    material = str(row[12]) if len(row) > 12 else ''
    return '黑碳晶' in color or '黑碳晶' in material


def extract_material_group(row):
    """按材质把行归到 多层 / 密度板 / 黑碳晶 / 复合 四类，不再按颜色分。"""
    color = str(row[7]) if len(row) > 7 else ''
    craft = str(row[11]) if len(row) > 11 else ''
    material = str(row[12]) if len(row) > 12 else ''

    if is_heitanjing(row):
        return '黑碳晶'
    if '密度板' in material:
        return '密度板'
    if '多层加密' in color or '多层加密' in craft:
        return '多层'
    return '复合'


def is_oversize_row(row, limit=OVERSIZE_LIMIT):
    try:
        return float(row[2]) > limit if len(row) > 2 and row[2] not in (None, '') else False
    except (ValueError, TypeError):
        return False


def build_oversize_meta(cat_type):
    if cat_type == 'menkuang':
        return {
            'type': 'menkuang',
            'base_color': '',
            'craft': '',
            'thickness': 0,
            'merge_suffix': f'超{OVERSIZE_LIMIT}',
            'name': f'门套超{OVERSIZE_LIMIT}',
            'allow_mixed_thickness': True,
        }
    return {
        'type': 'yakou',
        'base_color': '',
        'craft': '',
        'thickness': 0,
        'merge_suffix': f'超{OVERSIZE_LIMIT}',
        'name': f'窗套超{OVERSIZE_LIMIT}',
        'allow_mixed_thickness': True,
    }


def build_category_meta(cat_type, base_color, thickness, craft='', hidden=False, oversize=False):
    """Build a single source of truth for category identity.

    Category identity is product type + normalized color + craft + thickness
    (+ optional oversize flag). The returned name is only the display/output
    filename stem; all merge logic should prefer the structured fields so
    thickness cannot be dropped by a filename parsing edge case.
    """
    if hidden:
        hidden_thickness = 18
        name = f'{base_color}{craft}隐形门套{hidden_thickness}'
        return {
            'type': 'menkuang',
            'base_color': base_color,
            'craft': craft,
            'thickness': hidden_thickness,
            'merge_suffix': '隐形',
            'name': name,
        }

    oversize_suffix = f'超{OVERSIZE_LIMIT}' if oversize else ''

    if oversize:
        # 超长门套/窗套统一按材质命名，不再区分颜色/厚度
        if cat_type == 'menkuang':
            name = f'门套超长-{base_color}'
        elif cat_type == 'yakou':
            name = f'窗套超长-{base_color}'
        else:
            name = f'{base_color}{craft}{thickness_suffix}{oversize_suffix}'
        merge_suffix = f'超长-{base_color}'
        thickness_int = parse_thickness(thickness, 28 if cat_type == 'menkuang' else 18)
        return {
            'type': cat_type,
            'base_color': base_color,
            'craft': craft,
            'thickness': thickness_int,
            'merge_suffix': merge_suffix,
            'name': name,
        }

    if cat_type == 'menkuang':
        thickness_int = parse_thickness(thickness, 28)
        thickness_suffix = non_default_thickness_suffix(thickness_int, 28)
        name = f'{base_color}{craft}门套{thickness_suffix}'
    elif cat_type == 'yakou':
        thickness_int = parse_thickness(thickness, 18)
        thickness_suffix = non_default_thickness_suffix(thickness_int, 18)
        name = f'哑口套{base_color}{craft}{thickness_suffix}'
    elif cat_type == 'huqiang':
        thickness_int = parse_thickness(thickness, 18)
        thickness_suffix = f'厚度{thickness_int}'
        name = f'护墙-{base_color}{thickness_suffix}'
    else:
        thickness_int = parse_thickness(thickness, 0)
        thickness_suffix = non_default_thickness_suffix(thickness_int, 0)
        name = f'{base_color}{craft}{thickness_suffix}'

    # 非超尺寸行不加入 base_color，保证同一颜色族的不同颜色名仍能被合并。
    merge_suffix = f'{craft}{thickness_suffix}'
    return {
        'type': cat_type,
        'base_color': base_color,
        'craft': craft,
        'thickness': thickness_int,
        'merge_suffix': merge_suffix,
        'name': name,
    }


def add_category_row(categories, meta, row):
    cat_name = meta['name']
    if cat_name not in categories:
        categories[cat_name] = {'type': meta['type'], 'rows': [], 'meta': meta}
    categories[cat_name]['rows'].append(row)


def meta_from_category(cat_name, info):
    meta = info.get('meta') if isinstance(info, dict) else None
    if meta:
        return meta
    # Compatibility fallback for older ad-hoc category dictionaries.
    return {
        'type': info.get('type', '') if isinstance(info, dict) else '',
        'base_color': get_base_color_from_cat(cat_name),
        'craft': '多层加密' if '多层加密' in cat_name else '',
        'thickness': parse_thickness(re.search(r'(\d+)厚$', cat_name).group(1), 0) if re.search(r'(\d+)厚$', cat_name) else 0,
        'merge_suffix': get_cat_suffix(cat_name),
        'name': cat_name,
    }


def classify_menkuang_rows(rows, date_code, output_dir=None, reference_dir=None):
    """Classify 门框 rows into categories.
    
    隐形门套从门套中拆分出来，归入哑口套分类。
    同一个颜色既有哑口套又有隐形门套时，合并到同一个文件。
    """
    categories = {}
    for row in rows:
        item_name = row[0] if len(row) > 0 else ''
        color = row[7] if len(row) > 7 else ''
        thickness = row[5] if len(row) > 5 else 28
        base_color = extract_base_color(color)
        craft = craft_from_color(color)

        if '隐形门套' in str(item_name):
            # 隐形门套厚度强制为18
            if len(row) > 5:
                row[5] = 18
            meta = build_category_meta('menkuang', base_color, 18, craft, hidden=True)
        elif '门套' in str(item_name) and is_oversize_row(row):
            # 超过 2420 的不再按颜色分，只按材质（多层/密度板/黑碳晶/复合）分
            material = extract_material_group(row)
            craft = ''
            meta = build_category_meta('menkuang', material, thickness, craft, oversize=True)
        elif '门套' in str(item_name):
            if is_heitanjing(row):
                craft = '黑碳晶'
            meta = build_category_meta('menkuang', base_color, thickness, craft)
        else:
            meta = build_category_meta('menkuang', base_color, thickness, craft)

        add_category_row(categories, meta, row)

    for cat in categories.values():
        cat['rows'].sort(key=sort_key_for_row)

    return categories

def classify_yakou_rows(rows, date_code):
    """Classify 哑口套 rows into categories."""
    categories = {}
    for row in rows:
        color = row[7] if len(row) > 7 else ''
        thickness = row[5] if len(row) > 5 else 18
        base_color = extract_base_color(color)
        craft = craft_from_color(color)
        if is_oversize_row(row):
            # 超过 2420 的不再按颜色分，只按材质（多层/密度板/黑碳晶/复合）分
            material = extract_material_group(row)
            craft = ''
            meta = build_category_meta('yakou', material, thickness, craft, oversize=True)
        else:
            if is_heitanjing(row):
                craft = '黑碳晶'
            meta = build_category_meta('yakou', base_color, thickness, craft)
        add_category_row(categories, meta, row)

    for cat in categories.values():
        cat['rows'].sort(key=sort_key_for_row)

    return categories

def read_huqiang_data(xls_path):
    """Read 护墙 sheet (index 4), auto-detect column structure and skip red rows."""
    book = xlrd.open_workbook(xls_path, formatting_info=True)
    if len(book.sheet_names()) <= 4:
        return [], 0
    sheet = book.sheet_by_index(4)
    
    # Try to find header row
    header_row = 0
    col_idx = {}
    for r in range(min(5, sheet.nrows)):
        headers = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        for name in ['生产号', '颜色', '开料长', '开料宽', '数量', '厚度', '客户名称']:
            for c, h in enumerate(headers):
                if h == name:
                    col_idx[name] = c
        if col_idx:
            header_row = r
            break
    
    if not col_idx:
        return [], 0
    
    def get(row, name):
        c = col_idx.get(name)
        return sheet.cell_value(row, c) if c is not None else ''
    
    rows = []
    skipped = 0
    for r in range(header_row + 1, sheet.nrows):
        order_id = str(get(r, '生产号')).strip()
        if not order_id:
            continue
        # Check red font
        is_red = False
        for c in range(min(sheet.ncols, 12)):
            try:
                xf = sheet.cell_xf_index(r, c)
                fmt = book.xf_list[xf]
                font = book.font_list[fmt.font_index]
                if font.colour_index == 10:
                    is_red = True
                    break
            except Exception:
                pass
        if is_red:
            skipped += 1
            continue
        
        rows.append([
            order_id,
            get(r, '开料长'),
            get(r, '开料宽'),
            get(r, '数量'),
            get(r, '厚度'),
            get(r, '颜色'),
            get(r, '客户名称') if '客户名称' in col_idx else '',
        ])
    
    return rows, skipped

def classify_huqiang_rows(rows):
    """Classify 护墙 rows by color and thickness."""
    categories = {}
    for row in rows:
        color = str(row[5]) if len(row) > 5 else ''
        raw_thickness = row[4] if len(row) > 4 else ''
        try:
            thickness = int(float(raw_thickness)) if raw_thickness not in (None, '') else 18
        except (ValueError, TypeError):
            thickness = 18
        base_color = extract_base_color(color)
        if not base_color:
            continue
        meta = build_category_meta('huqiang', base_color, thickness)
        add_category_row(categories, meta, row)

    for cat in categories.values():
        cat['rows'].sort(key=lambda r: (str(r[0])))

    return categories

def get_base_color_from_cat(cat_name):
    # Strip the oversize suffix first so the existing regexes still work.
    name = re.sub(r'超\d+$', '', cat_name)
    # 超长类别直接返回材质名
    m = re.match(r'^门套超长-(.+)$', name)
    if m:
        return m.group(1)
    m = re.match(r'^窗套超长-(.+)$', name)
    if m:
        return m.group(1)
    # 非超长黑碳晶：颜色与材质分离
    m = re.match(r'^(.+?)黑碳晶门套(?:\d+厚)?$', name)
    if m:
        return m.group(1)
    m = re.match(r'^(.+?)门套(?:\d+厚)?$', name)
    if m:
        return m.group(1).replace('多层加密', '')
    m = re.match(r'^哑口套(.+?)黑碳晶(?:\d+厚)?$', name)
    if m:
        return m.group(1)
    m = re.match(r'^哑口套(.+)$', name)
    if m:
        return re.sub(r'\d+厚$', '', m.group(1).replace('多层加密', ''))
    m = re.match(r'^护墙-(.+)厚度\d+$', name)
    if m:
        return m.group(1).replace('多层加密', '')
    m = re.match(r'^(.+?)隐形门套\d+$', name)
    if m:
        return m.group(1).replace('多层加密', '')
    m = re.match(r'^隐形门套(.+)$', name)
    if m:
        return m.group(1).replace('多层加密', '')
    return name.replace('多层加密', '')

def get_cat_suffix(cat_name):
    """提取类别名的工艺/材质/厚度后缀，用于分组时保持工艺一致。"""
    suffix_parts = []
    if '隐形门套' in cat_name:
        suffix_parts.append('隐形')
        return ''.join(suffix_parts)
    # 材质组本身作为 suffix 的一部分，避免不同材质被合并
    if '黑碳晶' in cat_name:
        suffix_parts.append('黑碳晶')
    elif '密度板' in cat_name:
        suffix_parts.append('密度板')
    elif '多层' in cat_name:
        suffix_parts.append('多层')
    elif '复合' in cat_name:
        suffix_parts.append('复合')
    if '多层加密' in cat_name:
        m = re.search(r'(\d+厚)$', cat_name)
        suffix_parts.append(f'多层加密{m.group(1)}' if m else '多层加密')
    if '40厚' in cat_name:
        suffix_parts.append('40厚')
    if '15厚' in cat_name:
        suffix_parts.append('15厚')
    m = re.search(r'(超\d+)$', cat_name)
    if m:
        suffix_parts.append(m.group(1))
    return ''.join(suffix_parts)

def apply_color_merge(categories, output_dir, color_families, color_defaults=None, reference_dir=None):
    """Merge categories based on color families.
    
    同一颜色族、同类型、同工艺后缀的类别合并到一起，以数据量最多的类别名称作为文件名。
    如果数据量相同，优先保留有模板文件的类别；如果都没有模板，按类别名字母序。
    保持原始颜色名称，不进行自动重命名。
    """
    if not color_families or not output_dir or not os.path.exists(output_dir):
        return categories
    
    existing_files = set(os.listdir(output_dir))
    ref_files = set()
    if reference_dir and os.path.exists(reference_dir):
        ref_files = set(os.listdir(reference_dir))
    
    def file_exists(cat_name, cat_type):
        filename = f'{cat_name}.xls'
        return filename in existing_files or filename in ref_files
    
    def count_rows(info):
        total = 0
        for key in ['rows', 'rows_mk', 'rows_yk', 'rows_hq']:
            if key in info and info[key]:
                total += len(info[key])
        return total
    
    member_to_family = build_member_to_family(color_families)
    
    # Group categories by (color_family, type, suffix)
    # Key: (family_key, cat_type, suffix), Value: list of (cat_name, info)
    family_groups = {}
    
    for cat_name, info in categories.items():
        # Skip 隐形门套 - they should not be merged with regular categories
        if '隐形门套' in cat_name:
            continue
        meta = meta_from_category(cat_name, info)
        base_color = meta['base_color']
        norm_color = normalize_color_for_lookup(base_color)
        if norm_color in member_to_family:
            family_key = member_to_family[norm_color]
            suffix = meta['merge_suffix']
            group_key = (family_key, info['type'], suffix)
            if group_key not in family_groups:
                family_groups[group_key] = []
            family_groups[group_key].append((cat_name, info))
    
    merge_map = {}
    
    for group_key, members in family_groups.items():
        if len(members) <= 1:
            continue
        
        # Sort by: has_template (desc), row_count (desc), cat_name (asc for stability)
        def sort_key(item):
            cat_name, info = item
            has_template = 1 if file_exists(cat_name, info['type']) else 0
            row_count = count_rows(info)
            return (has_template, row_count, cat_name)
        
        members_sorted = sorted(members, key=sort_key, reverse=True)
        main_cat = members_sorted[0][0]
        
        for cat_name, info in members:
            if cat_name != main_cat:
                merge_map[cat_name] = main_cat
    
    # Apply merge
    merged = {}
    for cat_name, info in categories.items():
        if cat_name in merge_map:
            target = merge_map[cat_name]
            if target not in merged:
                merged[target] = {
                    'type': categories[target]['type'],
                    'meta': categories[target].get('meta'),
                    'rows_mk': [],
                    'rows_yk': [],
                    'rows_hq': [],
                }
            if 'rows_mk' in info:
                merged[target]['rows_mk'].extend(info['rows_mk'])
            if 'rows_yk' in info:
                merged[target]['rows_yk'].extend(info['rows_yk'])
            if 'rows' in info:
                if categories[target]['type'] == 'menkuang':
                    merged[target]['rows_mk'].extend(info['rows'])
                elif categories[target]['type'] == 'huqiang':
                    if 'rows_hq' not in merged[target]:
                        merged[target]['rows_hq'] = []
                    merged[target]['rows_hq'].extend(info['rows'])
                else:
                    merged[target]['rows_yk'].extend(info['rows'])
        else:
            if cat_name not in merged:
                merged[cat_name] = info
            else:
                if 'rows' in info:
                    if 'rows' not in merged[cat_name]:
                        merged[cat_name]['rows'] = []
                    merged[cat_name]['rows'].extend(info['rows'])
                if 'rows_mk' in info:
                    if 'rows_mk' not in merged[cat_name]:
                        merged[cat_name]['rows_mk'] = []
                    merged[cat_name]['rows_mk'].extend(info['rows_mk'])
                if 'rows_yk' in info:
                    if 'rows_yk' not in merged[cat_name]:
                        merged[cat_name]['rows_yk'] = []
                    merged[cat_name]['rows_yk'].extend(info['rows_yk'])
    
    for cat in merged.values():
        if 'rows' in cat:
            cat['rows'].sort(key=sort_key_for_row)
        if 'rows_mk' in cat:
            cat['rows_mk'].sort(key=sort_key_for_row)
        if 'rows_yk' in cat:
            cat['rows_yk'].sort(key=sort_key_for_row)
    
    return merged


def find_matching_yakou_for_hidden(yx_meta, all_cats, color_families, member_to_family):
    """Find an existing yakou category matching a hidden-door-frame family."""
    base_color = yx_meta['base_color']
    craft = yx_meta['craft']

    exact_name = build_category_meta('yakou', base_color, 18, craft)['name']
    if exact_name in all_cats and all_cats[exact_name].get('type') == 'yakou':
        return exact_name

    yx_family = member_to_family.get(normalize_color_for_lookup(base_color))
    if not yx_family:
        return None

    matches = []
    for cat_name, info in all_cats.items():
        if info.get('type') != 'yakou':
            continue
        meta = meta_from_category(cat_name, info)
        if meta['craft'] != craft or parse_thickness(meta['thickness'], 18) != 18:
            continue
        yakou_family = member_to_family.get(normalize_color_for_lookup(meta['base_color']))
        if yakou_family == yx_family:
            row_count = len(info.get('rows_yk', [])) + len(info.get('rows', []))
            matches.append((row_count, cat_name))

    if not matches:
        return None
    matches.sort(key=lambda item: (-item[0], item[1]))
    return matches[0][1]


def transform_row(row, idx, material, is_yinxing=False):
    """Transform original 18-col row to target 10-col format."""
    prefix = row[15] if len(row) > 15 else ''
    color = row[7] if len(row) > 7 else ''
    full_color = str(prefix) + str(color)
    
    gongyi_name = row[6] if len(row) > 6 else ''
    special_req = row[17] if len(row) > 17 else ''
    if special_req and '开合页孔' in str(special_req):
        gongyi_name = str(gongyi_name) + '/开合页孔'
    
    thickness = row[5] if len(row) > 5 else ''
    if is_yinxing and material == '垭口套':
        thickness = 18
    
    return [
        idx,
        material,
        row[2] if len(row) > 2 else '',
        row[3] if len(row) > 3 else '',
        row[4] if len(row) > 4 else '',
        thickness,
        gongyi_name,
        full_color,
        row[9] if len(row) > 9 else '',
        row[10] if len(row) > 10 else '',
    ]

def transform_huqiang_row(row, idx):
    """Transform 护墙 row [order,length,width,qty,thickness,color,customer] to 10-col."""
    return [
        idx,
        '护墙',
        row[1] if len(row) > 1 else '',
        row[2] if len(row) > 2 else '',
        row[3] if len(row) > 3 else '',
        row[4] if len(row) > 4 else '',
        '护墙',
        str(row[5]) if len(row) > 5 else '',
        str(row[0]) if len(row) > 0 else '',
        str(row[6]) if len(row) > 6 else '',
    ]

def to_number(value):
    """Best-effort numeric conversion for quantity totals."""
    if value in (None, ''):
        return 0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0

def format_number(value):
    return int(value) if isinstance(value, float) and value.is_integer() else value


def validate_single_thickness_category(cat_name, data_rows, allow_mixed=False):
    if allow_mixed:
        return
    thicknesses = {
        parse_thickness(row[5], 0)
        for row in data_rows
        if len(row) > 5 and row[5] not in (None, '')
    }
    thicknesses.discard(0)
    if len(thicknesses) > 1:
        ordered = ', '.join(str(v) for v in sorted(thicknesses))
        raise ValueError(f'分类 {cat_name} 混入多个厚度: {ordered}')


def get_template_file(output_dir, cat_type, existing_files, reference_dir=None, cat_name=None):
    """Find a suitable template file in output_dir or reference_dir."""
    is_oversize_cat = cat_name and '超长' in cat_name

    def find_template(files, directory):
        if cat_type == 'menkuang':
            # 优先找精确以“门套.xls”结尾的模板；找不到则放宽到任意“门套”xls
            candidates = [f for f in files
                          if f.endswith('门套.xls') and '40厚' not in f and '15厚' not in f and not f.startswith('哑口套') and not f.startswith('护墙')]
            if not candidates:
                candidates = [f for f in files
                              if '门套' in f and f.endswith('.xls') and not f.startswith('哑口套') and not f.startswith('护墙')]
            # 超长类别优先使用普通门套模板，避免匹配到自身超长文件
            if is_oversize_cat and candidates:
                normal_candidates = [f for f in candidates if '超长' not in f]
                if normal_candidates:
                    candidates = normal_candidates
        elif cat_type == 'yakou':
            candidates = [f for f in files if f.startswith('哑口套') and f.endswith('.xls')]
            if not candidates:
                candidates = [f for f in files if '哑口套' in f and f.endswith('.xls')]
            if is_oversize_cat and candidates:
                normal_candidates = [f for f in candidates if '超长' not in f]
                if normal_candidates:
                    candidates = normal_candidates
        elif cat_type == 'huqiang':
            candidates = [f for f in files if f.startswith('护墙') and f.endswith('.xls')]
        else:
            candidates = []
        if candidates:
            # 优先使用文件名最短的通用模板（如 门套.xls / 哑口套.xls）
            candidates.sort(key=len)
            return os.path.join(directory, candidates[0])
        return None
    
    template = find_template(existing_files, output_dir)
    if template:
        return template
    
    if reference_dir and os.path.exists(reference_dir):
        ref_files = [f for f in os.listdir(reference_dir) if f.endswith('.xls')]
        template = find_template(ref_files, reference_dir)
        if template:
            return template
    
    return None

def clone_cell_style_with_font_size(book, sheet, row_idx, col_idx, cache, font_points=12, num_format_str=None):
    """Clone an existing cell style and force the font size in points."""
    try:
        xf_index = sheet.cell_xf_index(row_idx, col_idx)
    except Exception:
        xf_index = 0

    key = (xf_index, font_points, num_format_str)
    if key in cache:
        return cache[key]

    rdxf = book.xf_list[xf_index]
    style = xlwt.XFStyle()
    style.num_format_str = num_format_str if num_format_str else book.format_map[rdxf.format_key].format_str

    src_font = book.font_list[rdxf.font_index]
    dst_font = style.font
    dst_font.height = font_points * 20
    dst_font.italic = src_font.italic
    dst_font.struck_out = src_font.struck_out
    dst_font.outline = src_font.outline
    dst_font.shadow = src_font.outline
    dst_font.colour_index = src_font.colour_index
    dst_font.bold = src_font.bold
    dst_font._weight = src_font.weight
    dst_font.escapement = src_font.escapement
    dst_font.underline = src_font.underline_type
    dst_font.family = src_font.family
    dst_font.charset = src_font.character_set
    dst_font.name = src_font.name

    style.protection.cell_locked = rdxf.protection.cell_locked
    style.protection.formula_hidden = rdxf.protection.formula_hidden

    style.borders.left = rdxf.border.left_line_style
    style.borders.right = rdxf.border.right_line_style
    style.borders.top = rdxf.border.top_line_style
    style.borders.bottom = rdxf.border.bottom_line_style
    style.borders.diag = rdxf.border.diag_line_style
    style.borders.left_colour = rdxf.border.left_colour_index
    style.borders.right_colour = rdxf.border.right_colour_index
    style.borders.top_colour = rdxf.border.top_colour_index
    style.borders.bottom_colour = rdxf.border.bottom_colour_index
    style.borders.diag_colour = rdxf.border.diag_colour_index
    style.borders.need_diag1 = rdxf.border.diag_down
    style.borders.need_diag2 = rdxf.border.diag_up

    style.pattern.pattern = rdxf.background.fill_pattern
    style.pattern.pattern_fore_colour = rdxf.background.pattern_colour_index
    style.pattern.pattern_back_colour = rdxf.background.background_colour_index

    style.alignment.horz = rdxf.alignment.hor_align
    style.alignment.vert = rdxf.alignment.vert_align
    style.alignment.dire = rdxf.alignment.text_direction
    style.alignment.rota = rdxf.alignment.rotation
    style.alignment.wrap = rdxf.alignment.text_wrapped
    style.alignment.shri = rdxf.alignment.shrink_to_fit
    style.alignment.inde = rdxf.alignment.indent_level

    cache[key] = style
    return style


def _biff_record(record_id, data):
    return struct.pack('<HH', record_id, len(data)) + data


def _find_biff_record(data, record_id):
    pos = 0
    data_len = len(data)
    while pos + 4 <= data_len:
        current_id, record_len = struct.unpack_from('<HH', data, pos)
        if current_id == record_id:
            return pos
        pos += 4 + record_len
    return -1


def _selection_record(first_row, last_row, col_idx):
    data = struct.pack(
        '<BHHHHHHBB',
        3,
        first_row,
        col_idx,
        0,
        1,
        first_row,
        last_row,
        col_idx,
        col_idx,
    )
    return _biff_record(0x001D, data)


def _autofilter_name_record(last_row, last_col):
    rpn = struct.pack('<BHHHHH', 0x3B, 0, 0, last_row, 0, last_col)
    data = (
        struct.pack('<HBBHHHBBBBB', 0x0021, 0, 1, len(rpn), 0, 1, 0, 0, 0, 0, 0)
        + b'\x0d'
        + rpn
    )
    return _biff_record(0x0018, data)


def _autofilter_records(col_count):
    drawing_group = bytes.fromhex(
        '0f0000f034000000000006f0180000000b080000020000000b0000000100000001'
        '0000000b00000023000bf00c000000810141000008c00140000008'
    )
    first_drawing = bytes.fromhex(
        '0f0002f090030000100008f0080000000b0000000a0400000f0003f078030000'
        '0f0004f028000000010009f010000000000000000000000000000000000000000'
        '2000af00800000000040000050000000f0004f04c000000920c0af008000000'
        '01040000000a000033000bf0120000007f0004010401bf0008000800bf030000'
        '0100000010f012000000010000000000000000000100000001000000000011f0'
        '00000000'
    )
    next_drawing_template = bytearray.fromhex(
        '0f0004f04c000000920c0af00800000002040000000a000033000bf012000000'
        '7f0004010401bf0008000800bf0300000100000010f012000000010001000000'
        '000000000200000001000000000011f000000000'
    )
    obj_template = bytearray.fromhex(
        '150012001400010001210000000000000000000000000c001400000000000000'
        '0000640001000a00000010000100130014000000000004000103000002000800'
        '4e0000000000'
    )

    col_count = max(1, min(col_count, 10))
    sheet_records = [_biff_record(0x009D, struct.pack('<H', col_count))]
    sheet_records.append(_biff_record(0x00EC, first_drawing))
    obj = bytearray(obj_template)
    obj[6:8] = struct.pack('<H', 1)
    sheet_records.append(_biff_record(0x005D, obj))

    for obj_id in range(2, col_count + 1):
        drawing = bytearray(next_drawing_template)
        drawing[16:20] = struct.pack('<I', 0x400 + obj_id)
        drawing[60:62] = struct.pack('<H', obj_id - 1)
        drawing[68:70] = struct.pack('<H', obj_id)
        sheet_records.append(_biff_record(0x00EC, drawing))

        obj = bytearray(obj_template)
        obj[6:8] = struct.pack('<H', obj_id)
        sheet_records.append(_biff_record(0x005D, obj))

    return _biff_record(0x00EB, drawing_group), b''.join(sheet_records)


def configure_excel_open_state(wb, ws, data_row_count, col_count=10, quantity_col_idx=4):
    """Add BIFF8 records for header filters and selecting the quantity column on open."""
    wb.active_sheet = 0
    wb.setup_ownbook()
    sheet_refs = getattr(wb, '_Workbook__sheet_refs')
    sheet_refs.clear()
    sheet_refs[(wb._ownbook_supbookx, 0, 0)] = 0

    col_count = max(1, min(col_count, 10))
    global_drawing_record, sheet_filter_records = _autofilter_records(col_count)
    filter_name_record = _autofilter_name_record(max(0, data_row_count), col_count - 1)
    original_sheet_get_biff_data = ws.get_biff_data

    def sheet_get_biff_data():
        data = original_sheet_get_biff_data()
        window2_pos = _find_biff_record(data, 0x023E)
        insert_pos = window2_pos if window2_pos >= 0 else max(0, len(data) - 4)
        data = data[:insert_pos] + sheet_filter_records + data[insert_pos:]

        eof_record = _biff_record(0x000A, b'')
        eof_pos = _find_biff_record(data, 0x000A)
        last_selected_row = max(1, data_row_count)
        selection = _selection_record(1, last_selected_row, quantity_col_idx)
        if eof_pos >= 0:
            data = data[:eof_pos] + selection + data[eof_pos:]
        return data

    ws.get_biff_data = sheet_get_biff_data

    def workbook_get_biff_data():
        before = b''
        for method_name in [
            '__bof_rec', '__intf_hdr_rec', '__intf_mms_rec', '__intf_end_rec',
            '__write_access_rec', '__codepage_rec', '__dsf_rec', '__tabid_rec',
            '__fngroupcount_rec', '__wnd_protect_rec', '__protect_rec',
            '__obj_protect_rec', '__password_rec', '__prot4rev_rec',
            '__prot4rev_pass_rec', '__backup_rec', '__hide_obj_rec',
            '__window1_rec', '__datemode_rec', '__precision_rec',
            '__refresh_all_rec', '__bookbool_rec',
            '__all_fonts_num_formats_xf_styles_rec', '__palette_rec',
            '__useselfs_rec',
        ]:
            before += getattr(wb, '_Workbook' + method_name)()

        country = wb._Workbook__country_rec()
        all_links = wb._Workbook__all_links_rec()
        shared_strings = wb._Workbook__sst_rec()
        after = country + global_drawing_record + all_links + filter_name_record + shared_strings
        ext_sst = wb._Workbook__ext_sst_rec(0)
        eof = wb._Workbook__eof_rec()

        worksheets = getattr(wb, '_Workbook__worksheets')
        worksheets[wb.active_sheet].selected = True
        sheets = b''
        sheet_biff_lens = []
        for sheet in worksheets:
            sheet_data = sheet.get_biff_data()
            sheets += sheet_data
            sheet_biff_lens.append(len(sheet_data))

        boundsheets = wb._Workbook__boundsheets_rec(
            len(before),
            len(after) + len(ext_sst) + len(eof),
            sheet_biff_lens,
        )
        sst_stream_pos = (
            len(before)
            + len(boundsheets)
            + len(country)
            + len(global_drawing_record)
            + len(all_links)
            + len(filter_name_record)
        )
        ext_sst = wb._Workbook__ext_sst_rec(sst_stream_pos)
        return before + boundsheets + after + ext_sst + eof + sheets

    wb.get_biff_data = workbook_get_biff_data


def write_to_template(target_path, data_rows, material, template_path=None):
    """Write data to target file, preserving format via xlutils (Linux-compatible)."""
    if not os.path.exists(target_path) and template_path and os.path.exists(template_path):
        shutil.copy2(template_path, target_path)

    if not os.path.exists(target_path):
        print(f"ERROR: Target file does not exist and no template available: {target_path}")
        return False

    try:
        book = xlrd.open_workbook(target_path, formatting_info=True)
        wb = copy(book)
        source_sheet = book.sheet_by_index(0)
        ws = wb.get_sheet(0)
        style_cache = {}

        # 覆盖写入 987 行（与原始 COM 逻辑保持一致，空白行用于清除旧数据）
        for i in range(987):
            if i < len(data_rows):
                row = data_rows[i]
            else:
                row = [i + 1, material, '', '', '', '', '', '', '', '']
            for c_idx, value in enumerate(row):
                # 序号、生产号、颜色等按字符串写入，防止日期自动转换
                if c_idx == 0:
                    # 条形码/序号列设置为数字格式
                    try:
                        value = int(float(value)) if value not in (None, '') else ''
                    except (ValueError, TypeError):
                        pass
                elif c_idx in (7, 8):
                    value = str(value)
                num_fmt = 'General' if c_idx == 0 else None
                style = clone_cell_style_with_font_size(book, source_sheet, i + 1, c_idx, style_cache, 12, num_fmt)
                ws.write(i + 1, c_idx, value, style)

        configure_excel_open_state(wb, ws, len(data_rows), 10, 4)
        wb.save(target_path)
        return True
    except Exception as e:
        print(f"ERROR writing to {target_path}: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_file(input_path, output_dir, reference_dir=None, color_map_path=None, skip_file=None):
    date_code = get_date_code_from_filename(input_path)
    if not date_code:
        print(f"Cannot extract date code from {input_path}")
        return
    
    print(f"Processing {input_path} -> {output_dir} (date: {date_code})")
    
    skip_set = load_skip_list(skip_file)
    if skip_set:
        print(f"  Skip list loaded: {len(skip_set)} order IDs")
    
    menkuang_rows = read_sheet_data(input_path, '门框')
    yakou_rows = read_sheet_data(input_path, '哑口套')
    
    menkuang_rows, mk_skipped = filter_skipped_rows(menkuang_rows, skip_set)
    yakou_rows, yk_skipped = filter_skipped_rows(yakou_rows, skip_set)
    
    all_skipped = set(mk_skipped + yk_skipped)
    if all_skipped:
        print(f"  Skipped {len(all_skipped)} order IDs")
        for oid in sorted(all_skipped):
            print(f"    - {oid}")
    
    print(f"  门框 rows: {len(menkuang_rows)}")
    print(f"  哑口套 rows: {len(yakou_rows)}")
    
    mk_cats = classify_menkuang_rows(menkuang_rows, date_code, output_dir, reference_dir)
    yk_cats = classify_yakou_rows(yakou_rows, date_code)
    hq_rows, hq_skipped = read_huqiang_data(input_path)
    if hq_rows or hq_skipped:
        print(f"  护墙 rows: {len(hq_rows)}" + (f" (skipped {hq_skipped})" if hq_skipped else ""))
    hq_cats = classify_huqiang_rows(hq_rows)
    color_families, color_defaults = load_color_families(color_map_path)
    member_to_family = build_member_to_family(color_families)
    
    all_cats = {}
    for name, info in mk_cats.items():
        all_cats[name] = {'type': info['type'], 'meta': info.get('meta'), 'rows_mk': info['rows'], 'rows_yk': [], 'rows_hq': []}
    for name, info in yk_cats.items():
        if name in all_cats:
            all_cats[name]['rows_yk'].extend(info['rows'])
            all_cats[name].setdefault('meta', info.get('meta'))
        else:
            all_cats[name] = {'type': info['type'], 'meta': info.get('meta'), 'rows_mk': [], 'rows_yk': info['rows'], 'rows_hq': []}
    for name, info in hq_cats.items():
        all_cats[name] = {'type': info['type'], 'meta': info.get('meta'), 'rows_hq': info['rows']}
    
    # 合并隐形门套到同颜色哑口套（如果存在），并保留多层加密等工艺后缀
    yinxing_names = [n for n in all_cats.keys() if '隐形门套' in n]
    for yx_name in yinxing_names:
        yx_meta = meta_from_category(yx_name, all_cats[yx_name])
        base_color = yx_meta['base_color']
        craft = yx_meta['craft']
        yakou_name = find_matching_yakou_for_hidden(yx_meta, all_cats, color_families, member_to_family)
        if yakou_name in all_cats:
            # 合并到哑口套分类
            if 'rows_mk' not in all_cats[yakou_name]:
                all_cats[yakou_name]['rows_mk'] = []
            all_cats[yakou_name]['rows_mk'].extend(all_cats[yx_name].get('rows_mk', []))
            del all_cats[yx_name]
            print(f"  Merged 隐形门套{base_color}{craft} -> {yakou_name}")
    
    if color_families:
        print(f"  Loaded color families: {len(color_families)} colors")
        os.makedirs(output_dir, exist_ok=True)
        all_cats = apply_color_merge(all_cats, output_dir, color_families, color_defaults, reference_dir)
    
    existing_files = [f for f in os.listdir(output_dir) if f.endswith('.xls')] if os.path.exists(output_dir) else []
    
    os.makedirs(output_dir, exist_ok=True)
    
    quantity_files = []
    quantity_total = 0.0

    for cat_name, cat_info in sorted(all_cats.items()):
        cat_type = cat_info['type']
        
        if cat_type == 'menkuang':
            material = '装门门框'
            filename = f'{cat_name}.xls'
        elif cat_type == 'huqiang':
            material = '护墙'
            filename = f'{cat_name}.xls'
        else:
            material = '垭口套'
            filename = f'{cat_name}.xls'
        
        target_path = os.path.join(output_dir, filename)
        
        if cat_type == 'huqiang':
            combined_rows = cat_info.get('rows_hq', [])
            data_rows = [transform_huqiang_row(r, i+1) for i, r in enumerate(combined_rows)]
        else:
            combined_rows = cat_info.get('rows_yk', []) + cat_info.get('rows_mk', [])
            data_rows = [transform_row(r, i+1, material) for i, r in enumerate(combined_rows)]
        meta = meta_from_category(cat_name, cat_info)
        validate_single_thickness_category(cat_name, data_rows, meta.get('allow_mixed_thickness', False))
        
        template_path = get_template_file(output_dir, cat_type, existing_files, reference_dir, cat_name)
        
        # If no template found and color defaults are available, try to find a family member's template
        if not template_path and color_defaults:
            base_color = meta['base_color']
            norm_color = normalize_color_for_lookup(base_color)
            if norm_color in color_defaults:
                # Find all related colors and their templates
                for related_color in color_families.get(norm_color, set()):
                    related_names = []
                    if cat_type == 'menkuang':
                        related_names.append(build_category_meta('menkuang', related_color, meta['thickness'], meta['craft'])['name'])
                        if meta['thickness'] != 28:
                            related_names.append(build_category_meta('menkuang', related_color, 28, meta['craft'])['name'])
                        if not meta['craft'] and meta['thickness'] == 28:
                            related_names.append(build_category_meta('menkuang', related_color, 18, '', hidden=True)['name'])
                    elif cat_type == 'yakou':
                        related_names.append(build_category_meta('yakou', related_color, meta['thickness'], meta['craft'])['name'])
                        if meta['thickness'] != 18:
                            related_names.append(build_category_meta('yakou', related_color, 18, meta['craft'])['name'])
                    elif cat_type == 'huqiang':
                        related_names.append(build_category_meta('huqiang', related_color, meta['thickness'])['name'])
                    
                    for related_name in related_names:
                        # Check output_dir first
                        related_path = os.path.join(output_dir, f'{related_name}.xls')
                        if os.path.exists(related_path):
                            template_path = related_path
                            break
                        # Then check reference_dir
                        if reference_dir:
                            related_path = os.path.join(reference_dir, f'{related_name}.xls')
                            if os.path.exists(related_path):
                                template_path = related_path
                                break
                    if template_path:
                        break
        
        print(f"  Writing {filename}: {len(data_rows)} rows (type={cat_type}, template={template_path})")
        success = write_to_template(target_path, data_rows, material, template_path)
        if success:
            file_total = sum(to_number(row[4] if len(row) > 4 else 0) for row in data_rows)
            quantity_total += file_total
            quantity_files.append({
                'filename': filename,
                'quantity_total': format_number(file_total),
            })
            print(f"    [OK]")
        else:
            print(f"    [FAIL]")
    
    print(f"Completed. Total categories: {len(all_cats)}")
    print(f"Quantity total: {format_number(quantity_total)}")
    return {
        'quantity_total': format_number(quantity_total),
        'quantity_files': quantity_files,
    }

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python classify_output.py <input_xls> <output_dir> [reference_dir] [color_map_json] [skip_list_txt]")
        print("")
        print("Parameters:")
        print("  input_xls      : Source .xls file with '门框' and '哑口套' sheets")
        print("  output_dir     : Target directory for classified files")
        print("  reference_dir  : (Optional) Directory to copy templates from if missing")
        print("  color_map_json : (Optional) Color family mapping JSON")
        print("  skip_list_txt  : (Optional) Text file with order IDs to skip (one per line, # for comments)")
        print("")
        print("Examples:")
        print("  python classify_output.py 'P5-24.xls' '5-24'")
        print("  python classify_output.py 'P6-03.xls' '6-03' '6-01' 'color_families.json' 'skip_orders.txt'")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_dir = sys.argv[2]
    reference_dir = sys.argv[3] if len(sys.argv) > 3 else None
    color_map_path = sys.argv[4] if len(sys.argv) > 4 else None
    skip_file = sys.argv[5] if len(sys.argv) > 5 else None
    
    process_file(input_path, output_dir, reference_dir, color_map_path, skip_file)
