#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
门套 sheet → 门框 sheet 转换脚本
用法: python transform_menkuang.py <输入xls文件路径> [输出xls文件路径]
保留原始"门框"sheet的格式，仅替换数据内容。
"""

import sys
import os
import shutil
import re
import xlrd
from xlutils.copy import copy


def extract_date_code(filepath):
    """从文件名提取日期代码。
    
    规则：
    - 文件名含 W 且在日期后：6-14W平板 -> W6-14
    - P 开头：P6-14 -> 6-14
    - 普通日期：6-14 -> 6-14
    - 旧格式：门套6-5 -> 6-5
    """
    basename = os.path.splitext(os.path.basename(filepath))[0]
    
    # W 在日期前：W6-14_for_classify.xls -> W6-14
    m = re.search(r'W(\d+-\d+)', basename, re.IGNORECASE)
    if m:
        return 'W' + m.group(1)
    
    # 日期后带 W：6-14W平板 / P6-14W平板 -> W6-14
    m = re.search(r'(\d+-\d+)[^\d-]*W', basename, re.IGNORECASE)
    if m:
        return 'W' + m.group(1)
    
    # P 开头：P6-14 -> 6-14
    m = re.match(r'^P(\d+-\d+)', basename)
    if m:
        return m.group(1)
    
    # 普通日期
    m = re.search(r'(\d+-\d+)', basename)
    if m:
        return m.group(1)
    
    # 旧格式：门套6-5
    m = re.search(r'(\d+-\d+)', basename.replace('门套', '').replace('窗套', ''))
    if m:
        return m.group(1)
    
    return "6-01"


def extract_template_params(book):
    """从原始输出sheet中提取模板参数"""
    params = {
        'prefix_base': None,
        'yinxing_suffix': None,
        'yinxing_col5_map': {},
        'jingyin_col6_suffix': None,
    }
    
    if '门框' not in book.sheet_names():
        return params
    
    ws = book.sheet_by_name('门框')
    
    # 提取前缀基础
    prefix_counts = {}
    for r in range(1, ws.nrows):
        p = ws.cell_value(r, 15) if ws.ncols > 15 else None
        if p:
            base = str(p)[1:] if str(p).startswith('急') else str(p)
            prefix_counts[base] = prefix_counts.get(base, 0) + 1
    if prefix_counts:
        params['prefix_base'] = max(prefix_counts, key=prefix_counts.get)
    
    # 提取隐形门套列6后缀
    suffix_counts = {}
    for r in range(1, ws.nrows):
        if ws.ncols > 9 and ws.cell_value(r, 0) == "隐形门套":
            col6 = str(ws.cell_value(r, 6))
            m = re.search(r'/隐形/(.*)', col6)
            if m:
                suffix_counts[m.group(1)] = suffix_counts.get(m.group(1), 0) + 1
    if suffix_counts:
        params['yinxing_suffix'] = max(suffix_counts, key=suffix_counts.get)
    
    # 提取隐形门套列5映射
    for r in range(1, ws.nrows):
        if ws.ncols > 9 and ws.cell_value(r, 0) == "隐形门套":
            pid = ws.cell_value(r, 9)
            params['yinxing_col5_map'][pid] = ws.cell_value(r, 5)
    
    # 提取静音门套列6后缀
    jingyin_suffixes = {}
    for r in range(1, ws.nrows):
        if ws.ncols > 6 and ws.cell_value(r, 0) == "静音门套":
            col6 = str(ws.cell_value(r, 6))
            # 提取"/静音"前面的部分，然后去掉单双前缀
            m = re.search(r'/(.*)', col6)
            if m:
                suffix = m.group(1)
                jingyin_suffixes[suffix] = jingyin_suffixes.get(suffix, 0) + 1
    if jingyin_suffixes:
        params['jingyin_col6_suffix'] = max(jingyin_suffixes, key=jingyin_suffixes.get)
    
    return params


def clean_color(color):
    """清理颜色中的重复模式"""
    s = str(color)
    if s.endswith('多层加密多层加密'):
        return s[:-4]
    if s.endswith('多层板加密多层板'):
        return s.replace('多层板加密多层板', '多层加密')
    return s


def extract_taoban_thickness(tsyao):
    """从特殊要求中提取套板厚度，如'套板厚40MM' -> 40"""
    m = re.search(r'套板厚(\d+)MM', str(tsyao))
    return int(m.group(1)) if m else None


def extract_jingyin_suffix(tsyao):
    """从特殊要求中提取静音门后缀"""
    s = str(tsyao)
    if s.startswith('静音门'):
        rest = s[3:].strip()
        if not rest:
            return '静音'
        # 取前两个"+"分隔的部分
        parts = rest.split('+')
        if len(parts) >= 2:
            return parts[0] + '+' + parts[1]
        return rest
    return '静音'


def build_col_map(ws):
    """读取第一行header，建立列名到索引的映射"""
    col_map = {}
    for c in range(ws.ncols):
        val = str(ws.cell_value(0, c)).strip()
        col_map[val] = c
    return col_map


def is_red_row(book, sheet, row_idx):
    """检查某行第一列字体是否为红色（colour_index == 10），标红表示'不做'"""
    try:
        xf = book.xf_list[sheet.cell_xf_index(row_idx, 0)]
        font = book.font_list[xf.font_index]
        return font.colour_index == 10
    except Exception:
        return False


def read_and_transform(input_path):
    """读取输入文件，转换数据，返回二维列表。支持新旧两种列结构。"""
    date_code = extract_date_code(input_path)
    book = xlrd.open_workbook(input_path, formatting_info=True)
    
    # 从原始输出中提取模板参数
    params = extract_template_params(book)
    prefix_base = date_code  # Always use date code from filename, not template prefix
    yinxing_suffix = params['yinxing_suffix'] or '见图生产'
    yinxing_col5_map = params['yinxing_col5_map']
    jingyin_col6_suffix = params['jingyin_col6_suffix']
    
    # 找到输入 sheet（包含 "门套" 的 sheet）
    src_sheet_name = None
    for name in book.sheet_names():
        if "门套" in name:
            src_sheet_name = name
            break
    if src_sheet_name is None:
        raise ValueError("输入文件中未找到包含'门套'的sheet")
    
    src_ws = book.sheet_by_name(src_sheet_name)
    col_map = build_col_map(src_ws)
    
    # 判断新旧格式：新格式有"厚度"列，旧格式没有
    is_new_format = '厚度' in col_map
    
    def get(row, key):
        idx = col_map.get(key)
        return row[idx] if idx is not None else None
    
    output_rows = []
    
    if is_new_format:
        # ===== 新格式：直接逐行转换，不再分组竖板/顶板 =====
        for r in range(1, src_ws.nrows):
            if is_red_row(book, src_ws, r):
                continue
            row_data = [src_ws.cell_value(r, c) for c in range(src_ws.ncols)]
            
            生产号 = get(row_data, '生产号')
            工件名称_raw = get(row_data, '工件名称')
            大类 = get(row_data, '大类')
            开料长 = get(row_data, '开料长')
            开料宽 = get(row_data, '开料宽')
            数量 = get(row_data, '数量')
            单双 = get(row_data, '单双')
            客户名称 = get(row_data, '客户名称')
            工艺 = get(row_data, '工艺')
            材质 = get(row_data, '材质')
            内填 = get(row_data, '内填')
            颜色 = clean_color(get(row_data, '颜色'))
            生产类型 = get(row_data, '生产类型')
            单项备注 = get(row_data, '单项备注')
            特殊要求 = get(row_data, '特殊要求')
            厚度 = get(row_data, '厚度')
            
            # 隐形门套只按“工件名称”列识别，特殊要求里出现“隐形”不再改类。
            ts_str = str(特殊要求) if 特殊要求 else ''
            item_name_text = str(工件名称_raw) if 工件名称_raw else ''
            is_yinxing = '隐形门套' in item_name_text
            # 静音芯板不等于静音门，仅当特殊要求含"静音门"时才视为静音门套
            is_jingyin = '静音门' in ts_str
            
            col6_base = str(单双).strip() if 单双 else "双面"
            
            # 单面时检查单项备注/特殊要求中的开槽信息
            col6_suffix = ""
            if '单面' in col6_base:
                remark_str = str(单项备注) if 单项备注 else ''
                special_str = str(特殊要求) if 特殊要求 else ''
                combined = remark_str + special_str
                for keyword in ['采台面开槽', '踩台面开槽', '合页面开槽']:
                    if keyword in combined:
                        col6_suffix = '/' + keyword
                        break
            
            if is_yinxing:
                工件名称 = '隐形门套'
                col5 = 厚度 if 厚度 else 18
                col6 = col6_base + col6_suffix + "/隐形/" + yinxing_suffix
            elif is_jingyin:
                工件名称 = '静音门套'
                col5 = 厚度 if 厚度 else 28
                suffix = jingyin_col6_suffix or extract_jingyin_suffix(特殊要求)
                col6 = col6_base + col6_suffix + "/" + suffix
            else:
                工件名称 = '门套'
                col5 = 厚度 if 厚度 else 28
                col6 = col6_base + col6_suffix
            
            # 前缀规则
            if 生产类型 == "正常生产":
                prefix = prefix_base
            else:
                prefix = "急" + prefix_base
            
            col8 = prefix + str(颜色)
            col15 = prefix
            
            output_rows.append([
                工件名称,
                大类 if 大类 else '门套',
                开料长,
                开料宽,
                数量,
                col5,
                col6,
                颜色,
                col8,
                生产号,
                客户名称,
                工艺,
                材质,
                内填,
                生产类型,
                col15,
                单项备注,
                特殊要求,
            ])
    else:
        # ===== 旧格式：按生产号分组竖板、顶板 =====
        groups = {}  # 生产号 -> {"竖板": [rows], "顶板": [rows]}
        for r in range(1, src_ws.nrows):
            if is_red_row(book, src_ws, r):
                continue
            row_data = [src_ws.cell_value(r, c) for c in range(src_ws.ncols)]
            pid = get(row_data, '生产号')
            category = get(row_data, '大类')
            if pid not in groups:
                groups[pid] = {"竖板": [], "顶板": []}
            if category in groups[pid]:
                groups[pid][category].append(row_data)
        
        # 按生产号字符串升序排序
        sorted_groups = sorted(groups.items(), key=lambda x: str(x[0]))
        
        for pid, group in sorted_groups:
            shuban_list = group.get("竖板", [])
            dingban_list = group.get("顶板", [])
            
            if not shuban_list or not dingban_list:
                print(f"警告: 生产号 {pid} 缺少竖板或顶板，跳过")
                continue
            
            for row_data in shuban_list + dingban_list:
                生产号 = get(row_data, '生产号')
                工件名称 = get(row_data, '工件名称')
                大类 = get(row_data, '大类')
                开料长 = get(row_data, '开料长')
                开料宽 = get(row_data, '开料宽')
                数量 = get(row_data, '数量')
                单双 = get(row_data, '单双')
                客户名称 = get(row_data, '客户名称')
                工艺 = get(row_data, '工艺')
                材质 = get(row_data, '材质')
                内填 = get(row_data, '内填')
                颜色 = clean_color(get(row_data, '颜色'))
                生产类型 = get(row_data, '生产类型')
                单项备注 = get(row_data, '单项备注')
                特殊要求 = get(row_data, '特殊要求')
                
                # 列5 / 列6 规则
                taoban_thickness = extract_taoban_thickness(特殊要求)
                
                if taoban_thickness:
                    col5 = taoban_thickness
                elif 工件名称 == "隐形门套":
                    col5 = yinxing_col5_map.get(生产号, 18)
                elif 工件名称 == "静音门套":
                    col5 = 28
                else:
                    col5 = 28
                
                if 工件名称 == "隐形门套":
                    col6 = str(单双) + "/隐形/" + yinxing_suffix
                elif 工件名称 == "静音门套":
                    suffix = jingyin_col6_suffix or extract_jingyin_suffix(特殊要求)
                    col6 = str(单双) + "/" + suffix
                else:
                    col6 = "双面"
                
                # 前缀规则
                if 生产类型 == "正常生产":
                    prefix = prefix_base
                else:
                    prefix = "急" + prefix_base
                
                col8 = prefix + str(颜色)
                col15 = prefix
                
                output_rows.append([
                    工件名称,
                    大类,
                    开料长,
                    开料宽,
                    数量,
                    col5,
                    col6,
                    颜色,
                    col8,
                    生产号,
                    客户名称,
                    工艺,
                    材质,
                    内填,
                    生产类型,
                    col15,
                    单项备注,
                    特殊要求,
                ])
    
    return output_rows


def write_with_format(template_path, output_path, data_rows):
    """使用 xlutils 保留模板格式并写入数据（兼容 Linux Docker）"""
    book = xlrd.open_workbook(template_path, formatting_info=True)
    wb = copy(book)

    try:
        sheet_name = "门框"
        ws = wb.get_sheet(sheet_name)
    except Exception:
        sheet_name = book.sheet_by_index(0).name
        ws = wb.get_sheet(0)
    source_ws = book.sheet_by_name(sheet_name)

    # 覆盖写入数据行，从第 2 行开始
    total_cols = len(data_rows[0]) if data_rows else max(source_ws.ncols, 18)
    if data_rows:
        for r_idx, row in enumerate(data_rows):
            for c_idx in range(total_cols):
                value = row[c_idx] if c_idx < len(row) else ''
                # 序号、前缀、生产号等列按字符串写入，避免被解析为日期
                if c_idx in (0, 8, 9, 15):
                    value = str(value)
                ws.write(r_idx + 1, c_idx, value)

    # 清掉模板残留旧数据，避免 P6-01 模板尾部行被后续分类当作本次输入。
    for r_idx in range(len(data_rows) + 1, source_ws.nrows):
        for c_idx in range(total_cols):
            ws.write(r_idx, c_idx, '')

    wb.save(output_path)


def transform(input_path, output_path=None):
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = base + "_output" + ext
    
    print(f"读取输入: {input_path}")
    data_rows = read_and_transform(input_path)
    print(f"转换完成: {len(data_rows)} 行数据")
    
    print(f"写入输出（保留格式）: {output_path}")
    write_with_format(input_path, output_path, data_rows)
    print("完成!")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python transform_menkuang.py <输入xls文件路径> [输出xls文件路径]")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    transform(input_path, output_path)
