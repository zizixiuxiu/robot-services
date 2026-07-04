#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
经销商数据报表生成脚本
从"账号指标"和"设计师数据统计"两个数据源生成格式统一的经销商报表。
支持两种格式：
  - old / 旧版：与原 3月/4月/6.2 经销商数据表格式一致
  - new / 新版：与 6月经销商酷家乐数据(新) 格式一致，含大区合计、总合计、美化样式
"""

import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill


# ==================== 配置项 ====================

DEFAULT_REGION_ORDER = ['东北区', '华北区', '华东区', '华南区', '华中区', '西北区', '西南区']

# 新版格式颜色方案
COLOR_TITLE_BG = '1F4E78'      # 深蓝
COLOR_HEADER_BG = '4472C4'     # 蓝色
COLOR_REGION_TOTAL_BG = 'D9E1F2'  # 浅蓝
COLOR_GRAND_TOTAL_BG = '1F4E78'   # 深蓝（总合计，与标题呼应）
COLOR_ZEBRA = 'F2F2F2'         # 浅灰
COLOR_WHITE = 'FFFFFF'
COLOR_BORDER = 'BFBFBF'

# 字体样式
FONT_TITLE = Font(name='黑体', size=18, bold=True, color='FFFFFF')
FONT_HEADER = Font(name='黑体', size=11, bold=True, color='FFFFFF')
FONT_DATA = Font(name='微软雅黑', size=10)
FONT_TOTAL = Font(name='微软雅黑', size=11, bold=True)
FONT_TOTAL_BLUE = Font(name='微软雅黑', size=11, bold=True, color='1F4E78')
FONT_TOTAL_WHITE = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')

# 边框样式
THIN_BORDER = Border(
    left=Side(style='thin', color='FF' + COLOR_BORDER),
    right=Side(style='thin', color='FF' + COLOR_BORDER),
    top=Side(style='thin', color='FF' + COLOR_BORDER),
    bottom=Side(style='thin', color='FF' + COLOR_BORDER)
)

# 居中对齐
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)


# ==================== 填充样式对象 ====================

def fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type='solid')


# ==================== 数据清洗函数 ====================

def clean_dept6_name(row):
    """
    清洗数据源中的经销商名称。
    优先使用"部门6"，为空时回退到"账号名称"。
    去掉末尾的数字标记，如：
      '冯海东3/4' -> '冯海东'
      '张守芹4'   -> '张守芹'
      '张晋芳 2'  -> '张晋芳'
    """
    dept6 = row['部门6']
    account = row['账号名称']
    if pd.notna(dept6):
        s = str(dept6).strip()
    else:
        s = str(account).strip() if pd.notna(account) else ''

    s = re.sub(r'\s+\d+/\d+$', '', s)
    s = re.sub(r'\s+\d+$', '', s)
    s = re.sub(r'\d+/\d+$', '', s)
    s = re.sub(r'\d+$', '', s)
    s = s.strip()
    return s


def clean_designer_dept6(s):
    """
    清洗设计师数据中的部门名称最后一级。
    过滤掉纯区域名称（如"华东区+上海..."）和空值。
    """
    if pd.isna(s):
        return ''
    s = str(s).strip()
    if re.match(r'^[\u4e00-\u9fa5]{2}区\+', s) or s == '':
        return ''

    s = re.sub(r'\s+\d+/\d+$', '', s)
    s = re.sub(r'\s+\d+$', '', s)
    s = re.sub(r'\d+/\d+$', '', s)
    s = re.sub(r'\d+$', '', s)
    s = s.strip()
    return s


# ==================== 核心处理逻辑 ====================

def process_account_data(account_file, sheet_name='账号信息明细'):
    """
    读取账号指标数据，清洗并聚合。
    返回 DataFrame：['区域', '经销商', '创建方案数', '新增渲染方案数']

    过滤规则：
    - 只保留"部门5"包含"+"号的行（即正常的大区+省份格式），
      这样可以排除内部部门（如直营店、研发部、设计支持部等）
      以及其他非经销商账号。
    """
    df = pd.read_excel(account_file, sheet_name=sheet_name)

    df = df[df['部门5'].astype(str).str.contains(r'\+', na=False)].copy()

    df['经销商'] = df.apply(clean_dept6_name, axis=1)
    df['区域'] = df['部门5'].str.extract(r'^([^+]+)')

    agg = df.groupby(['区域', '经销商']).agg({
        '创建方案数': 'sum',
        '新增渲染方案数': 'sum'
    }).reset_index()

    agg['区域_order'] = agg['区域'].apply(
        lambda x: DEFAULT_REGION_ORDER.index(x) if x in DEFAULT_REGION_ORDER else 999
    )
    agg = agg.sort_values(['区域_order', '经销商']).drop(
        columns='区域_order'
    ).reset_index(drop=True)

    return agg


def process_designer_data(designer_file, sheet_name='设计师数据统计'):
    """
    读取设计师数据统计，提取并聚合提审方案数。
    返回 DataFrame：['经销商', '提审方案数']
    """
    df = pd.read_excel(designer_file, sheet_name=sheet_name)

    df['部门6_raw'] = df['部门名称'].str.split('/', n=4).str[-1]
    df['经销商'] = df['部门6_raw'].apply(clean_designer_dept6)
    df = df[df['经销商'] != '']

    agg = df.groupby('经销商').agg({'提审方案数': 'sum'}).reset_index()
    return agg


def build_output_df(account_agg, designer_agg):
    """
    合并账号数据和设计师数据，构建最终输出 DataFrame。
    """
    out = pd.DataFrame()
    out['区域'] = account_agg['区域']
    out['经销商'] = account_agg['经销商']
    out['创建方案数'] = account_agg['创建方案数']
    out['新增渲染方案数'] = account_agg['新增渲染方案数']
    out['方案深化占比'] = ''
    out['提审方案数'] = ''

    out = pd.merge(out, designer_agg, on='经销商', how='left')
    out['提审方案数'] = out['提审方案数_y'].fillna(0).astype(int)
    out = out.drop(columns=['提审方案数_x', '提审方案数_y'])

    return out


# ==================== 通用写入辅助 ====================

def _write_header(ws, title, headers, max_col):
    """写入标题行和表头行"""
    # 标题行：A1 写标题，其他列写空字符串但统一设置填充/边框，
    # 不合并单元格，避免 openpyxl 合并后部分列格式丢失
    title_fill = fill(COLOR_TITLE_BG)
    for col in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col, value=title if col == 1 else '')
        cell.font = FONT_TITLE
        cell.alignment = CENTER_ALIGN
        cell.fill = title_fill
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 55.0

    # 列标题
    header_fill = fill(COLOR_HEADER_BG)
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col)
        cell.value = header
        cell.font = FONT_HEADER
        cell.alignment = CENTER_ALIGN
        cell.fill = header_fill
        cell.border = THIN_BORDER
    ws.row_dimensions[2].height = 40.0


def _set_column_widths(ws, widths):
    """设置列宽"""
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


# ==================== 旧版格式写入 ====================

def write_formatted_excel_old(df, output_path, title='经销商数据'):
    """
    旧版格式：与原 3月/4月/6.2 经销商数据表格式一致
    """
    wb = Workbook()
    ws = wb.active
    ws.title = '账号信息明细'

    _set_column_widths(ws, {
        'A': 17.875, 'B': 32.875, 'C': 21.75,
        'D': 11.0, 'E': 16.25, 'F': 16.75
    })

    headers = ['区域', '经销商', '创建方案数', '新增渲染方案数', '方案深化占比', '提审方案数']
    _write_header(ws, title, headers, 6)

    # 数据行
    for r_idx, row in enumerate(df.itertuples(index=False), 3):
        cell_a = ws.cell(row=r_idx, column=1, value=row.区域 if row.区域 else None)
        cell_a.font = FONT_DATA
        cell_a.alignment = CENTER_ALIGN
        cell_a.border = THIN_BORDER

        ws.cell(row=r_idx, column=2, value=row.经销商).font = FONT_DATA
        ws.cell(row=r_idx, column=2).alignment = CENTER_ALIGN
        ws.cell(row=r_idx, column=2).border = THIN_BORDER

        ws.cell(row=r_idx, column=3, value=row.创建方案数).font = FONT_DATA
        ws.cell(row=r_idx, column=3).alignment = CENTER_ALIGN
        ws.cell(row=r_idx, column=3).border = THIN_BORDER

        ws.cell(row=r_idx, column=4, value=row.新增渲染方案数).font = FONT_DATA
        ws.cell(row=r_idx, column=4).alignment = CENTER_ALIGN
        ws.cell(row=r_idx, column=4).border = THIN_BORDER

        if row.创建方案数 > 0:
            cell_e = ws.cell(row=r_idx, column=5, value=f'=D{r_idx}/C{r_idx}')
        else:
            cell_e = ws.cell(row=r_idx, column=5, value=0)
        cell_e.font = FONT_DATA
        cell_e.alignment = CENTER_ALIGN
        cell_e.border = THIN_BORDER
        cell_e.number_format = '0%'

        ws.cell(row=r_idx, column=6, value=row.提审方案数).font = FONT_DATA
        ws.cell(row=r_idx, column=6).alignment = CENTER_ALIGN
        ws.cell(row=r_idx, column=6).border = THIN_BORDER

        ws.row_dimensions[r_idx].height = 17.0

    # 合并区域列
    current_region = None
    start_row = 3
    for r_idx, region in enumerate(df['区域'], 3):
        if region != '' and region != current_region:
            if current_region is not None:
                ws.merge_cells(start_row=start_row, start_column=1,
                               end_row=r_idx - 1, end_column=1)
            current_region = region
            start_row = r_idx
    if current_region is not None:
        ws.merge_cells(start_row=start_row, start_column=1,
                       end_row=len(df) + 2, end_column=1)

    wb.save(output_path)
    print(f'旧版报表已生成: {output_path}')
    print(f'共 {len(df)} 行经销商数据')


# ==================== 新版格式写入（含美化） ====================

def write_formatted_excel_new(df, output_path, title='经销商数据'):
    """
    新版格式：与 6月经销商酷家乐数据(新) 格式一致
    包含：大区合计、总合计、提审/渲染、提审/创建，并带美化样式
    """
    wb = Workbook()
    ws = wb.active
    ws.title = '账号信息明细'

    _set_column_widths(ws, {
        'A': 24, 'B': 30, 'C': 14, 'D': 16,
        'E': 16, 'F': 16, 'G': 12, 'H': 12
    })

    headers = [
        '区域', '经销商', '创建方案数',
        '新增渲染方案数\n(出效果图)',
        '方案深化占比\n(渲染/创建)',
        '提审方案数\n(下单到工厂)',
        '提审/渲染', '提审/创建'
    ]
    _write_header(ws, title, headers, 8)

    # 构建区域块
    region_blocks = []
    current_region = None
    current_rows = []

    for idx, region in enumerate(df['区域']):
        if region != '' and region != current_region:
            if current_region is not None:
                region_blocks.append((current_region, current_rows))
            current_region = region
            current_rows = []
        current_rows.append(idx)
    if current_region is not None:
        region_blocks.append((current_region, current_rows))

    data_rows_info = []
    current_excel_row = 3

    for region_name, row_indices in region_blocks:
        block_start = current_excel_row
        data_idx_in_block = 0

        for data_idx in row_indices:
            row = df.iloc[data_idx]
            data_idx_in_block += 1
            data_rows_info.append({
                'excel_row': current_excel_row,
                'data_idx': data_idx,
                'is_data': True
            })

            # 斑马纹：隔行底色
            row_fill = fill(COLOR_ZEBRA) if data_idx_in_block % 2 == 0 else fill(COLOR_WHITE)

            cell_a = ws.cell(row=current_excel_row, column=1,
                             value=region_name if data_idx == row_indices[0] else None)
            cell_a.fill = row_fill
            cell_a.font = FONT_DATA
            cell_a.alignment = CENTER_ALIGN
            cell_a.border = THIN_BORDER

            for col, val in [(2, row['经销商']), (3, row['创建方案数']),
                             (4, row['新增渲染方案数']), (6, row['提审方案数'])]:
                cell = ws.cell(row=current_excel_row, column=col, value=val)
                cell.fill = row_fill
                cell.font = FONT_DATA
                cell.alignment = CENTER_ALIGN
                cell.border = THIN_BORDER

            # 方案深化占比公式
            if row['创建方案数'] > 0:
                cell_e = ws.cell(row=current_excel_row, column=5,
                                 value=f'=D{current_excel_row}/C{current_excel_row}')
            else:
                cell_e = ws.cell(row=current_excel_row, column=5, value=0)
            cell_e.fill = row_fill
            cell_e.font = FONT_DATA
            cell_e.alignment = CENTER_ALIGN
            cell_e.border = THIN_BORDER
            cell_e.number_format = '0%'

            # 提审/渲染：提审和渲染都为0时留空
            if row['新增渲染方案数'] > 0 and row['提审方案数'] > 0:
                cell_g = ws.cell(row=current_excel_row, column=7,
                                 value=f'=F{current_excel_row}/D{current_excel_row}')
                cell_g.number_format = '0.00%'
            else:
                cell_g = ws.cell(row=current_excel_row, column=7, value='')
            cell_g.fill = row_fill
            cell_g.font = FONT_DATA
            cell_g.alignment = CENTER_ALIGN
            cell_g.border = THIN_BORDER

            # 提审/创建：提审和创建都为0时留空
            if row['创建方案数'] > 0 and row['提审方案数'] > 0:
                cell_h = ws.cell(row=current_excel_row, column=8,
                                 value=f'=F{current_excel_row}/C{current_excel_row}')
                cell_h.number_format = '0.00%'
            else:
                cell_h = ws.cell(row=current_excel_row, column=8, value='')
            cell_h.fill = row_fill
            cell_h.font = FONT_DATA
            cell_h.alignment = CENTER_ALIGN
            cell_h.border = THIN_BORDER

            ws.row_dimensions[current_excel_row].height = 22.0
            current_excel_row += 1

        # 大区合计行
        total_row = current_excel_row
        sum_range_c = f'C{block_start}:C{total_row - 1}'
        sum_range_d = f'D{block_start}:D{total_row - 1}'
        sum_range_f = f'F{block_start}:F{total_row - 1}'

        total_fill = fill(COLOR_REGION_TOTAL_BG)

        for col in range(1, 9):
            ws.cell(row=total_row, column=col).fill = total_fill
            ws.cell(row=total_row, column=col).font = FONT_TOTAL_BLUE
            ws.cell(row=total_row, column=col).alignment = CENTER_ALIGN
            ws.cell(row=total_row, column=col).border = THIN_BORDER

        ws.cell(row=total_row, column=2, value='大区合计')
        ws.cell(row=total_row, column=3, value=f'=SUM({sum_range_c})')
        ws.cell(row=total_row, column=4, value=f'=SUM({sum_range_d})')
        ws.cell(row=total_row, column=5, value=f'=D{total_row}/C{total_row}')
        ws.cell(row=total_row, column=5).number_format = '0.00%'
        ws.cell(row=total_row, column=6, value=f'=SUM({sum_range_f})')
        ws.cell(row=total_row, column=7, value=f'=F{total_row}/D{total_row}')
        ws.cell(row=total_row, column=7).number_format = '0.00%'
        ws.cell(row=total_row, column=8, value=f'=F{total_row}/C{total_row}')
        ws.cell(row=total_row, column=8).number_format = '0.00%'

        ws.row_dimensions[total_row].height = 28.0
        current_excel_row += 1

    # 空行
    empty_row = current_excel_row
    for col in range(1, 9):
        ws.cell(row=empty_row, column=col).border = THIN_BORDER
    ws.row_dimensions[empty_row].height = 12.0
    current_excel_row += 1

    # 总合计行
    grand_total_row = current_excel_row
    total_create = int(df['创建方案数'].sum())
    total_render = int(df['新增渲染方案数'].sum())
    total_submit = int(df['提审方案数'].sum())

    total_fill = fill(COLOR_GRAND_TOTAL_BG)

    for col in range(1, 9):
        ws.cell(row=grand_total_row, column=col).fill = total_fill
        ws.cell(row=grand_total_row, column=col).font = FONT_TOTAL_WHITE
        ws.cell(row=grand_total_row, column=col).alignment = CENTER_ALIGN
        ws.cell(row=grand_total_row, column=col).border = THIN_BORDER

    ws.cell(row=grand_total_row, column=2, value='总合计')
    ws.cell(row=grand_total_row, column=3, value=total_create)
    ws.cell(row=grand_total_row, column=4, value=total_render)
    cell_e = ws.cell(row=grand_total_row, column=5,
                     value=f'=D{grand_total_row}/C{grand_total_row}')
    cell_e.number_format = '0.00%'
    ws.cell(row=grand_total_row, column=6, value=total_submit)
    cell_g = ws.cell(row=grand_total_row, column=7,
                     value=f'=F{grand_total_row}/D{grand_total_row}')
    cell_g.number_format = '0.00%'
    cell_h = ws.cell(row=grand_total_row, column=8,
                     value=f'=F{grand_total_row}/C{grand_total_row}')
    cell_h.number_format = '0.00%'

    ws.row_dimensions[grand_total_row].height = 45.0

    # 合并每个区域的A列
    for region_name, row_indices in region_blocks:
        first_row = data_rows_info[row_indices[0]]['excel_row']
        last_row = data_rows_info[row_indices[-1]]['excel_row']
        ws.merge_cells(start_row=first_row, start_column=1,
                       end_row=last_row, end_column=1)

    # 冻结窗格：冻结前两行
    ws.freeze_panes = 'A3'

    wb.save(output_path)
    print(f'新版美化报表已生成: {output_path}')
    print(f'共 {len(df)} 行经销商数据，{len(region_blocks)} 个大区')


# ==================== 主函数 ====================

def main():
    parser = argparse.ArgumentParser(
        description='从账号指标和设计师数据生成经销商报表'
    )
    parser.add_argument(
        '--account',
        default=(r'D:\wechat\xwechat_files\wxid_0fh4oxng8dq212_f810\msg\file\2026-06\账号指标20260602.xlsx'),
        help='账号指标Excel文件路径'
    )
    parser.add_argument(
        '--designer',
        default=(r'D:\wechat\xwechat_files\wxid_0fh4oxng8dq212_f810\msg\file\2026-06\设计师数据统计 .xlsx'),
        help='设计师数据统计Excel文件路径'
    )
    parser.add_argument(
        '--output',
        default=(r'D:\wechat\xwechat_files\wxid_0fh4oxng8dq212_f810\msg\file\2026-06\经销商数据.xlsx'),
        help='输出报表路径'
    )
    parser.add_argument(
        '--title',
        default='6月经销商数据',
        help='报表标题（第一行显示内容）'
    )
    parser.add_argument(
        '--format',
        choices=['old', 'new'],
        default='new',
        help='输出格式：old=旧版（无合计），new=新版（含大区合计、总合计、美化样式）'
    )
    args = parser.parse_args()

    for path, name in [(args.account, '账号指标'), (args.designer, '设计师数据')]:
        if not Path(path).exists():
            raise FileNotFoundError(f'{name}文件不存在: {path}')

    account_agg = process_account_data(args.account)
    designer_agg = process_designer_data(args.designer)
    out_df = build_output_df(account_agg, designer_agg)

    if args.format == 'old':
        write_formatted_excel_old(out_df, args.output, title=args.title)
    else:
        write_formatted_excel_new(out_df, args.output, title=args.title)


if __name__ == '__main__':
    main()
