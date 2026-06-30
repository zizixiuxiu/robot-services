#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
窗套 sheet → 哑口套 sheet 转换脚本
用法: python transform_window_yakou.py <输入xls文件路径> [输出xls文件路径]
保留原始"哑口套"sheet的格式，仅替换数据内容。
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


def clean_color(color):
    """清理颜色中的重复模式"""
    s = str(color)
    if s.endswith('多层加密多层加密'):
        return s[:-4]
    if s.endswith('多层板加密多层板'):
        return s.replace('多层板加密多层板', '多层加密')
    return s


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
    
    src_sheet_name = None
    for name in book.sheet_names():
        if name == "窗套":
            src_sheet_name = name
            break
    if src_sheet_name is None:
        raise ValueError("输入文件中未找到'窗套'sheet")
    
    src_ws = book.sheet_by_name(src_sheet_name)
    col_map = build_col_map(src_ws)
    
    # 判断新旧格式：新格式有"厚度"列，旧格式没有
    is_new_format = '厚度' in col_map
    
    def get(row, key):
        idx = col_map.get(key)
        return row[idx] if idx is not None else None
    
    output_rows = []
    
    if is_new_format:
        # ===== 新格式：直接逐行转换 =====
        for r in range(1, src_ws.nrows):
            if is_red_row(book, src_ws, r):
                continue
            row_data = [src_ws.cell_value(r, c) for c in range(src_ws.ncols)]
            
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
            
            # 列5：窗套→哑口套保留原始厚度，15厚需后续单独分表
            raw_thickness = get(row_data, '厚度')
            try:
                col5 = int(float(raw_thickness)) if raw_thickness not in (None, '') else 18
            except (ValueError, TypeError):
                col5 = 18
            
            # 列6：直接复制输入单双值
            col6 = str(单双).strip() if 单双 else ""
            
            # 前缀规则
            if 生产类型 == "正常生产":
                prefix = date_code
            else:
                prefix = "急" + date_code
            
            col8 = prefix + str(颜色)
            col15 = prefix
            
            output_rows.append([
                工件名称,
                大类 if 大类 else '窗套',
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
            
            if not shuban_list and not dingban_list:
                print(f"警告: 生产号 {pid} 无数据，跳过")
                continue
            
            # 判断该生产号是否有多个顶板（用于添加"/对接"）
            has_multi_dingban = len(dingban_list) > 1
            
            # 先全部竖板，再全部顶板（与输入顺序一致）
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
                
                # 列5：旧格式窗套→哑口套固定为18（旧格式无独立厚度列）
                col5 = 18
                
                # 列6：直接复制输入单双值
                col6 = str(单双) if 单双 else ""
                # 如果输入中没有"/对接"，但该生产号有多个顶板，则添加"/对接"
                if 大类 == "顶板" and "/对接" not in col6 and has_multi_dingban:
                    col6 = col6 + "/对接"
                
                # 前缀规则
                if 生产类型 == "正常生产":
                    prefix = date_code
                else:
                    prefix = "急" + date_code
                
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
        sheet_name = "哑口套"
        ws = wb.get_sheet(sheet_name)
    except Exception:
        sheet_name = book.sheet_by_index(0).name
        ws = wb.get_sheet(0)
    source_ws = book.sheet_by_name(sheet_name)

    total_cols = len(data_rows[0]) if data_rows else max(source_ws.ncols, 18)
    if data_rows:
        for r_idx, row in enumerate(data_rows):
            for c_idx in range(total_cols):
                value = row[c_idx] if c_idx < len(row) else ''
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
        output_path = base + "_yakou_output" + ext
    
    print(f"读取输入: {input_path}")
    data_rows = read_and_transform(input_path)
    print(f"转换完成: {len(data_rows)} 行数据")
    
    print(f"写入输出（保留格式）: {output_path}")
    write_with_format(input_path, output_path, data_rows)
    print("完成!")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python transform_window_yakou.py <输入xls文件路径> [输出xls文件路径]")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    transform(input_path, output_path)
