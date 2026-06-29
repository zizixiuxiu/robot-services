#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单文件自动处理脚本：原始 xls -> 中间文件（过滤红色行）-> transform -> classify
兼容 Linux Docker，不再依赖 Windows COM 或子进程调用。
"""

import os
import sys
import shutil
import re
import tempfile
import xlrd
import xlwt
from pathlib import Path

# 脚本所在目录作为基础目录
SCRIPT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(SCRIPT_DIR))

import classify_output
import transform_menkuang
import transform_window_yakou

TEMPLATE_PATH = SCRIPT_DIR / 'test_P6-01_output3.xls'
REF_DIR = SCRIPT_DIR / 'ref_classification'
COLOR_MAP = SCRIPT_DIR / 'color_families.json'
BASE_OUTPUT_DIR = Path(os.getenv('OUTPUT_BASE', str(SCRIPT_DIR.parent / 'data' / 'output')))


def is_red_row(book, sheet, row_idx):
    """检查某行第一列字体是否为红色（colour_index == 10），标红表示'不做'"""
    try:
        xf = book.xf_list[sheet.cell_xf_index(row_idx, 0)]
        font = book.font_list[xf.font_index]
        return font.colour_index == 10
    except Exception:
        return False


def extract_date_code(filepath):
    """从文件名提取日期代码。

    规则：
    - 文件名含 W 且在日期后：6-14W平板 -> W6-14
    - P 开头：P6-14 -> 6-14
    - 普通日期：6-14 -> 6-14
    - 旧格式：门套6-5 -> 6-5
    """
    basename = os.path.splitext(os.path.basename(filepath))[0]

    # W 在日期前：W6-14.xls -> W6-14
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

    return None


def copy_sheet_to_writer(book, writer, sheet_name, new_name=None, skip_red=False):
    """复制 sheet 到 writer，可选跳过红色行"""
    old_sheet = book.sheet_by_name(sheet_name)
    new_sheet = writer.add_sheet(new_name or sheet_name)
    out_row = 0
    blank_style = xlwt.XFStyle()
    for row in range(old_sheet.nrows):
        if skip_red and row > 0 and is_red_row(book, old_sheet, row):
            continue
        for col in range(old_sheet.ncols):
            value = old_sheet.cell_value(row, col)
            new_sheet.write(out_row, col, value, blank_style)
        out_row += 1
    return new_sheet


def run_pipeline(src_path, final_output):
    """构建中间文件并运行 transform"""
    final_output = Path(final_output)
    temp_output = final_output.with_suffix('')
    temp_output = Path(str(temp_output) + '_temp.xls')

    # 1. 复制模板
    shutil.copy(str(TEMPLATE_PATH), str(temp_output))

    # 2. 读取源文件
    book_src = xlrd.open_workbook(src_path, formatting_info=True)

    menkuang_sheet_name = None
    chuangtao_sheet_name = None
    has_huqiang = False
    for name in book_src.sheet_names():
        if '门套' in name and '窗套' not in name:
            menkuang_sheet_name = name
        if name == '窗套':
            chuangtao_sheet_name = name
        if name == '护墙':
            has_huqiang = True

    if not menkuang_sheet_name:
        raise ValueError(f"未找到门套 sheet in {src_path}")
    if not chuangtao_sheet_name:
        raise ValueError(f"未找到窗套 sheet in {src_path}")

    mk_sheet_src = book_src.sheet_by_name(menkuang_sheet_name)
    ct_sheet_src = book_src.sheet_by_name(chuangtao_sheet_name)

    # 统计红色行
    mk_red = sum(1 for r in range(1, mk_sheet_src.nrows) if is_red_row(book_src, mk_sheet_src, r))
    ct_red = sum(1 for r in range(1, ct_sheet_src.nrows) if is_red_row(book_src, ct_sheet_src, r))
    if mk_red or ct_red:
        print(f'  [Filter] 过滤红色行: 门套={mk_red}, 窗套={ct_red}')

    # 3. 读取目标模板
    book_dst = xlrd.open_workbook(str(temp_output), formatting_info=True)
    writer = xlwt.Workbook(style_compression=2)

    for sheet_name in book_dst.sheet_names():
        if sheet_name in ['门框', '哑口套', '护墙']:
            continue
        if '门套' in sheet_name or '窗套' in sheet_name:
            continue
        copy_sheet_to_writer(book_dst, writer, sheet_name)

    # 复制源数据 sheet，跳过红色行
    copy_sheet_to_writer(book_src, writer, menkuang_sheet_name, '门套', skip_red=True)
    copy_sheet_to_writer(book_src, writer, chuangtao_sheet_name, '窗套', skip_red=True)

    for sheet_name in ['门框', '哑口套']:
        if sheet_name in book_dst.sheet_names():
            copy_sheet_to_writer(book_dst, writer, sheet_name)

    if has_huqiang:
        copy_sheet_to_writer(book_src, writer, '护墙', '护墙', skip_red=True)

    writer.save(str(temp_output))
    print(f'  [Pipeline] 中间文件: {temp_output}')
    print(f'  [Pipeline] 门套: {mk_sheet_src.nrows} rows -> {mk_sheet_src.nrows - mk_red}, 窗套: {ct_sheet_src.nrows} rows -> {ct_sheet_src.nrows - ct_red}')

    # 4. transform_menkuang
    print("  [Pipeline] Running transform_menkuang...")
    transform_menkuang.transform(str(temp_output), str(final_output))

    # 5. transform_window_yakou
    print("  [Pipeline] Running transform_window_yakou...")
    yakou_path = str(final_output) + '.yakou'
    transform_window_yakou.transform(str(final_output), yakou_path)

    if os.path.exists(yakou_path):
        shutil.copy2(yakou_path, str(final_output))
        os.remove(yakou_path)

    os.remove(str(temp_output))
    print(f"  [Pipeline] 转换完成: {final_output}")
    return str(final_output)


def process_input(src_path, dir_name=None, output_base=None, debug_output=False):
    """处理单个输入文件，返回生成的所有输出文件路径列表。

    默认在临时目录中完成处理（HTTP 模式）；debug_output=True 时复制到 OUTPUT_BASE。
    """
    src_path = Path(src_path)
    if not src_path.exists():
        raise FileNotFoundError(f"文件不存在: {src_path}")

    if not dir_name:
        dir_name = extract_date_code(str(src_path))
        if not dir_name:
            dir_name = src_path.stem

    date_code = extract_date_code(str(src_path)) or dir_name

    work_dir = Path(tempfile.mkdtemp(prefix='pvc_'))
    try:
        final_output = work_dir / f'{date_code}_for_classify.xls'
        print(f"{'='*60}")
        print(f"处理文件: {src_path}")
        print(f"工作目录: {work_dir}")
        print(f"{'='*60}")

        # 1. Pipeline
        run_pipeline(str(src_path), str(final_output))

        # 2. Classify
        auto_dir = work_dir / f'{dir_name}-自动分类'
        if auto_dir.exists():
            shutil.rmtree(auto_dir)
        auto_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n  [Classify] Running classify_output...")
        classify_output.process_file(str(final_output), str(auto_dir), str(REF_DIR), str(COLOR_MAP))

        output_files = sorted([f for f in auto_dir.rglob('*.xls') if f.is_file()])
        result = {
            'work_dir': str(work_dir),
            'auto_dir': str(auto_dir),
            'output_files': [str(f) for f in output_files],
            'count': len(output_files),
        }

        # 调试模式：复制到 OUTPUT_BASE
        if debug_output:
            out_base = Path(output_base) if output_base else BASE_OUTPUT_DIR
            debug_dir = out_base / f'{dir_name}-自动分类'
            if debug_dir.exists():
                shutil.rmtree(debug_dir)
            shutil.copytree(auto_dir, debug_dir)
            result['debug_dir'] = str(debug_dir)
            print(f"  [Debug] 已复制到: {debug_dir}")

        print(f"\n{'='*60}")
        print(f"完成！输出: {auto_dir}")
        print(f"文件数: {len(output_files)}")
        print(f"{'='*60}")
        return result
    except Exception:
        # 异常时保留工作目录便于排查；正常流程由调用方清理
        import traceback
        traceback.print_exc()
        raise


def main():
    if len(sys.argv) < 2:
        print("用法: python process_single.py <源文件路径> [输出目录名]")
        print("示例: python process_single.py \"D:\\Edge下载\\P6-6.xls\"")
        sys.exit(1)

    src_path = sys.argv[1]
    if not os.path.exists(src_path):
        print(f"错误: 文件不存在 {src_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        dir_name = sys.argv[2]
    else:
        dir_name = None

    process_input(src_path, dir_name=dir_name, debug_output=True)


if __name__ == '__main__':
    main()
