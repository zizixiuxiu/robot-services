#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
经销商数据报表生成脚本
从"账号指标"和"设计师数据统计"两个数据源生成格式统一的经销商报表。
"""

import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side


# ==================== 配置项 ====================

DEFAULT_REGION_ORDER = ['东北区', '华北区', '华东区', '华南区', '华中区', '西北区', '西南区']

# 列宽（与原文件保持一致）
COLUMN_WIDTHS = {
    'A': 17.875,
    'B': 32.875,
    'C': 21.75,
    'D': 11.0,
    'E': 16.25,
    'F': 16.75,
}

# 字体样式
FONT_TITLE = Font(name='黑体', size=16)
FONT_HEADER = Font(name='黑体', size=12)
FONT_DATA = Font(name='黑体', size=10)

# 边框样式
THIN_BORDER = Border(
    left=Side(style='thin', color='FF000000'),
    right=Side(style='thin', color='FF000000'),
    top=Side(style='thin', color='FF000000'),
    bottom=Side(style='thin', color='FF000000')
)

# 居中对齐
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')


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

    # 去掉末尾的空格+数字标记
    s = re.sub(r'\s+\d+/\d+$', '', s)
    s = re.sub(r'\s+\d+$', '', s)
    # 去掉末尾的无空格数字标记
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

    # 过滤：部门5必须包含"+"，表示是正常的"区域+省份"格式
    df = df[df['部门5'].astype(str).str.contains(r'\+', na=False)].copy()

    df['经销商'] = df.apply(clean_dept6_name, axis=1)
    df['区域'] = df['部门5'].str.extract(r'^([^+]+)')

    agg = df.groupby(['区域', '经销商']).agg({
        '创建方案数': 'sum',
        '新增渲染方案数': 'sum'
    }).reset_index()

    # 按区域顺序排序
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

    # 部门名称格式：奢匠事业部/.../区域+省份/部门6
    # 用 split('/', n=4) 避免部门6中的"/"（如"2/3"）被误分割
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
    # 同区域只保留第一行显示，后续为空（用于Excel合并单元格效果）
    out.loc[out['区域'] == out['区域'].shift(1), '区域'] = ''
    out['经销商'] = account_agg['经销商']
    out['创建方案数'] = account_agg['创建方案数']
    out['新增渲染方案数'] = account_agg['新增渲染方案数']
    out['方案深化占比'] = ''  # 占位，后面写公式
    out['提审方案数'] = ''    # 占位，后面合并

    # 合并提审方案数
    out = pd.merge(out, designer_agg, on='经销商', how='left')
    out['提审方案数'] = out['提审方案数_y'].fillna(0).astype(int)
    out = out.drop(columns=['提审方案数_x', '提审方案数_y'])

    return out


# ==================== Excel 写入与格式化 ====================

def write_formatted_excel(df, output_path, title='经销商数据'):
    """
    将 DataFrame 写入 Excel，并保持与原文件一致的格式：
    - 黑体字体、指定字号
    - 全表细边框
    - 列宽、行高
    - 区域列合并单元格
    - 方案深化占比使用公式 =D/C，格式 0%
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = '账号信息明细'

    # 设置列宽
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # ---- 第1行：大标题 ----
    ws['A1'] = title
    ws.merge_cells('A1:F1')
    ws['A1'].font = FONT_TITLE
    ws['A1'].alignment = CENTER_ALIGN
    ws.row_dimensions[1].height = 51.0
    for col in range(1, 7):
        ws.cell(row=1, column=col).border = THIN_BORDER

    # ---- 第2行：列标题 ----
    headers = ['区域', '经销商', '创建方案数', '新增渲染方案数', '方案深化占比', '提审方案数']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = FONT_HEADER
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER
    ws.row_dimensions[2].height = 37.0

    # ---- 数据行 ----
    for r_idx, row in enumerate(df.itertuples(index=False), 3):
        # A: 区域
        cell_a = ws.cell(row=r_idx, column=1, value=row.区域 if row.区域 else None)
        cell_a.font = FONT_DATA
        cell_a.alignment = CENTER_ALIGN
        cell_a.border = THIN_BORDER

        # B: 经销商
        cell_b = ws.cell(row=r_idx, column=2, value=row.经销商)
        cell_b.font = FONT_DATA
        cell_b.alignment = CENTER_ALIGN
        cell_b.border = THIN_BORDER

        # C: 创建方案数
        cell_c = ws.cell(row=r_idx, column=3, value=row.创建方案数)
        cell_c.font = FONT_DATA
        cell_c.alignment = CENTER_ALIGN
        cell_c.border = THIN_BORDER

        # D: 新增渲染方案数
        cell_d = ws.cell(row=r_idx, column=4, value=row.新增渲染方案数)
        cell_d.font = FONT_DATA
        cell_d.alignment = CENTER_ALIGN
        cell_d.border = THIN_BORDER

        # E: 方案深化占比（公式）
        if row.创建方案数 > 0:
            cell_e = ws.cell(row=r_idx, column=5, value=f'=D{r_idx}/C{r_idx}')
        else:
            cell_e = ws.cell(row=r_idx, column=5, value=0)
        cell_e.font = FONT_DATA
        cell_e.alignment = CENTER_ALIGN
        cell_e.border = THIN_BORDER
        cell_e.number_format = '0%'

        # F: 提审方案数
        cell_f = ws.cell(row=r_idx, column=6, value=row.提审方案数)
        cell_f.font = FONT_DATA
        cell_f.alignment = CENTER_ALIGN
        cell_f.border = THIN_BORDER

        ws.row_dimensions[r_idx].height = 17.0

    # ---- 合并区域列 ----
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

    # 保存
    wb.save(output_path)
    print(f'报表已生成: {output_path}')
    print(f'共 {len(df)} 行经销商数据')


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
    args = parser.parse_args()

    # 检查输入文件
    for path, name in [(args.account, '账号指标'), (args.designer, '设计师数据')]:
        if not Path(path).exists():
            raise FileNotFoundError(f'{name}文件不存在: {path}')

    # 处理数据
    account_agg = process_account_data(args.account)
    designer_agg = process_designer_data(args.designer)
    out_df = build_output_df(account_agg, designer_agg)

    # 写入Excel
    write_formatted_excel(out_df, args.output, title=args.title)


if __name__ == '__main__':
    main()
