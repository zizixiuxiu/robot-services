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
import xlrd
from xlutils.copy import copy

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
    c = str(color).strip().upper()
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

def sort_key_for_row(row):
    """Sort key: normal production first, then by order ID."""
    production_type = row[14] if len(row) > 14 else ''
    order_id = row[9] if len(row) > 9 else ''
    is_urgent = 1 if '加急' in str(production_type) else 0
    return (is_urgent, str(order_id))

def classify_menkuang_rows(rows, date_code, output_dir=None, reference_dir=None):
    """Classify 门框 rows into categories.
    
    隐形门套从门套中拆分出来，归入哑口套分类。
    同一个颜色既有哑口套又有隐形门套时，合并到同一个文件。
    """
    categories = {}
    for row in rows:
        color = row[7] if len(row) > 7 else ''
        thickness = row[5] if len(row) > 5 else 28
        gongyi = row[6] if len(row) > 6 else ''
        prefix = row[15] if len(row) > 15 else date_code
        base_color = extract_base_color(color)
        
        if '隐形' in str(gongyi):
            # 隐形门套先单独分类，后续在 process_file 中决定是否合并到同颜色哑口套
            # 保留多层加密等工艺后缀，避免与普通哑口套错误合并
            suffix = '多层加密' if '多层加密' in str(color) else ''
            cat_name = f'{base_color}{suffix}隐形门套18'
            cat_type = 'menkuang'
            # 隐形门套厚度强制为18
            if len(row) > 5:
                row[5] = 18
        else:
            if '多层加密' in str(color):
                cat_name = f'{base_color}多层加密门套'
                cat_type = 'menkuang'
            elif thickness == 40:
                cat_name = f'{base_color}门套40厚'
                cat_type = 'menkuang'
            elif thickness == 15:
                cat_name = f'{base_color}门套15厚'
                cat_type = 'menkuang'
            else:
                cat_name = f'{base_color}门套'
                cat_type = 'menkuang'
        
        if cat_name not in categories:
            categories[cat_name] = {'type': cat_type, 'rows': []}
        categories[cat_name]['rows'].append(row)
    
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
        if '多层加密' in str(color):
            cat_name = f'哑口套{base_color}多层加密'
        elif thickness == 15:
            cat_name = f'哑口套{base_color}15厚'
        else:
            cat_name = f'哑口套{base_color}'
        if cat_name not in categories:
            categories[cat_name] = {'type': 'yakou', 'rows': []}
        categories[cat_name]['rows'].append(row)
    
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
        cat_name = f'护墙-{base_color}厚度{int(thickness)}'
        if not base_color:
            continue
        if cat_name not in categories:
            categories[cat_name] = {'type': 'huqiang', 'rows': []}
        categories[cat_name]['rows'].append(row)

    for cat in categories.values():
        cat['rows'].sort(key=lambda r: (str(r[0])))

    return categories

def get_base_color_from_cat(cat_name):
    m = re.match(r'^(.+?)门套(?:40厚|15厚)?$', cat_name)
    if m:
        return m.group(1).replace('多层加密', '')
    m = re.match(r'^哑口套(.+)$', cat_name)
    if m:
        return m.group(1).replace('多层加密', '').replace('15厚', '')
    m = re.match(r'^护墙-(.+)厚度\d+$', cat_name)
    if m:
        return m.group(1).replace('多层加密', '')
    m = re.match(r'^(.+?)隐形门套\d+$', cat_name)
    if m:
        return m.group(1).replace('多层加密', '')
    m = re.match(r'^隐形门套(.+)$', cat_name)
    if m:
        return m.group(1).replace('多层加密', '')
    return cat_name.replace('多层加密', '')

def get_cat_suffix(cat_name):
    """提取类别名的工艺后缀，用于分组时保持工艺一致。"""
    if '隐形门套' in cat_name:
        return '隐形'
    if '多层加密' in cat_name:
        return '多层加密'
    if '40厚' in cat_name:
        return '40厚'
    if '15厚' in cat_name:
        return '15厚'
    return ''

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
    
    # Build a mapping from each member color to its family key
    # This handles cases where the same family has multiple keys (e.g., ZKY5023 and YSM8897)
    member_to_family = {}
    for family_key, members in color_families.items():
        for member in members:
            norm_member = normalize_color_for_lookup(member)
            if norm_member not in member_to_family:
                member_to_family[norm_member] = family_key
    
    # Group categories by (color_family, type, suffix)
    # Key: (family_key, cat_type, suffix), Value: list of (cat_name, info)
    family_groups = {}
    
    for cat_name, info in categories.items():
        # Skip 隐形门套 - they should not be merged with regular categories
        if '隐形门套' in cat_name:
            continue
        base_color = get_base_color_from_cat(cat_name)
        norm_color = normalize_color_for_lookup(base_color)
        if norm_color in member_to_family:
            family_key = member_to_family[norm_color]
            suffix = get_cat_suffix(cat_name)
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
                merged[target] = {'type': categories[target]['type'], 'rows_mk': [], 'rows_yk': []}
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

def get_template_file(output_dir, cat_type, existing_files, reference_dir=None):
    """Find a suitable template file in output_dir or reference_dir."""
    def find_template(files, directory):
        if cat_type == 'menkuang':
            candidates = [f for f in files 
                          if f.endswith('门套.xls') and '40厚' not in f and '15厚' not in f and not f.startswith('哑口套') and not f.startswith('护墙')]
        elif cat_type == 'yakou':
            candidates = [f for f in files if f.startswith('哑口套') and f.endswith('.xls')]
        elif cat_type == 'huqiang':
            candidates = [f for f in files if f.startswith('护墙') and f.endswith('.xls')]
        else:
            candidates = []
        if candidates:
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
        ws = wb.get_sheet(0)

        # 覆盖写入 987 行（与原始 COM 逻辑保持一致，空白行用于清除旧数据）
        for i in range(987):
            if i < len(data_rows):
                row = data_rows[i]
            else:
                row = [i + 1, material, '', '', '', '', '', '', '', '']
            for c_idx, value in enumerate(row):
                # 序号、生产号、颜色等按字符串写入，防止日期自动转换
                if c_idx in (0, 7, 8):
                    value = str(value)
                ws.write(i + 1, c_idx, value)

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
    
    all_cats = {}
    for name, info in mk_cats.items():
        all_cats[name] = {'type': info['type'], 'rows_mk': info['rows'], 'rows_yk': [], 'rows_hq': []}
    for name, info in yk_cats.items():
        if name in all_cats:
            all_cats[name]['rows_yk'].extend(info['rows'])
        else:
            all_cats[name] = {'type': info['type'], 'rows_mk': [], 'rows_yk': info['rows'], 'rows_hq': []}
    for name, info in hq_cats.items():
        all_cats[name] = {'type': info['type'], 'rows_hq': info['rows']}
    
    # 合并隐形门套到同颜色哑口套（如果存在），并保留多层加密等工艺后缀
    yinxing_names = [n for n in all_cats.keys() if '隐形门套' in n]
    for yx_name in yinxing_names:
        base_color = get_base_color_from_cat(yx_name)
        suffix = '多层加密' if '多层加密' in yx_name else ''
        yakou_name = f'哑口套{base_color}{suffix}'
        if yakou_name in all_cats:
            # 合并到哑口套分类
            if 'rows_mk' not in all_cats[yakou_name]:
                all_cats[yakou_name]['rows_mk'] = []
            all_cats[yakou_name]['rows_mk'].extend(all_cats[yx_name].get('rows_mk', []))
            del all_cats[yx_name]
            print(f"  Merged 隐形门套{base_color}{suffix} -> {yakou_name}")
    
    color_families, color_defaults = load_color_families(color_map_path)
    if color_families:
        print(f"  Loaded color families: {len(color_families)} colors")
        os.makedirs(output_dir, exist_ok=True)
        all_cats = apply_color_merge(all_cats, output_dir, color_families, color_defaults, reference_dir)
    
    existing_files = [f for f in os.listdir(output_dir) if f.endswith('.xls')] if os.path.exists(output_dir) else []
    
    os.makedirs(output_dir, exist_ok=True)
    
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
        
        template_path = get_template_file(output_dir, cat_type, existing_files, reference_dir)
        
        # If no template found and color defaults are available, try to find a family member's template
        if not template_path and color_defaults:
            base_color = get_base_color_from_cat(cat_name)
            norm_color = normalize_color_for_lookup(base_color)
            if norm_color in color_defaults:
                # Find all related colors and their templates
                for related_color in color_families.get(norm_color, set()):
                    # Build potential category names for this related color
                    if cat_type == 'menkuang':
                        if '40厚' in cat_name:
                            related_names = [f'{related_color}门套40厚']
                        elif '15厚' in cat_name:
                            related_names = [f'{related_color}门套15厚']
                        else:
                            related_names = [f'{related_color}门套', f'{related_color}隐形门套18']
                    elif cat_type == 'yakou':
                        if '15厚' in cat_name:
                            related_names = [f'哑口套{related_color}15厚']
                        else:
                            related_names = [f'哑口套{related_color}']
                    elif cat_type == 'huqiang':
                        m = re.search(r'厚度(\d+)', cat_name)
                        thick = m.group(1) if m else '18'
                        related_names = [f'护墙-{related_color}厚度{thick}']
                    else:
                        related_names = []
                    
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
            print(f"    [OK]")
        else:
            print(f"    [FAIL]")
    
    print(f"Completed. Total categories: {len(all_cats)}")

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
