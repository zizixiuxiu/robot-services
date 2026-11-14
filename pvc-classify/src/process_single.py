#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单文件自动处理脚本：原始 xls -> 中间文件（过滤红色行）-> transform -> classify
兼容 Linux Docker，不再依赖 Windows COM 或子进程调用。
"""

import os
import sys
import json
import shutil
import re
import csv
import uuid
import tempfile
import urllib.request
import xlrd
import xlwt
from pathlib import Path
from keep_rows import (is_keep_row, read_row_values, reset_events, get_events,
                       format_events, add_excluded, add_kept)

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

# 宿主机 Excel 重写服务（CSV 来源的输出必须经 Excel/WPS 重写后南兴才能导入）
RESAVE_URL = os.getenv('PVC_RESAVE_URL', 'http://host.docker.internal:8017/resave')
RESAVE_EXCHANGE_CONTAINER = Path(os.getenv('PVC_RESAVE_DIR', '/app/data/resave_exchange'))
RESAVE_EXCHANGE_HOST = os.getenv('PVC_RESAVE_HOST_DIR', r'D:\Services\robot-services\pvc-classify\data\resave_exchange')


def resave_via_host(dst_dir):
    """把 dst_dir 下的 xls 经宿主机 Excel COM 重写（打开另存），南兴只认重写后的文件。
    通过 bind mount 的 resave_exchange 目录交换文件；失败时抛异常，不交付未重写的文件。"""
    xls_files = [f for f in os.listdir(dst_dir) if f.lower().endswith('.xls')]
    if not xls_files:
        return
    exchange_id = uuid.uuid4().hex[:12]
    c_dir = RESAVE_EXCHANGE_CONTAINER / exchange_id
    c_dir.mkdir(parents=True, exist_ok=True)
    try:
        for f in xls_files:
            shutil.copy2(os.path.join(dst_dir, f), c_dir / f)
        host_dir = str(Path(RESAVE_EXCHANGE_HOST) / exchange_id)
        req = urllib.request.Request(
            RESAVE_URL,
            data=json.dumps({'dir': host_dir}).encode('utf-8'),
            headers={'Content-Type': 'application/json'},
        )
        with urllib.request.urlopen(req, timeout=300) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
        if not payload.get('success'):
            raise RuntimeError(f"Excel 重写服务返回失败: {payload.get('error')}")
        print(f"  [Resave] 宿主机 Excel 重写完成: {payload.get('count')} 个文件")
        for f in xls_files:
            shutil.copy2(c_dir / f, os.path.join(dst_dir, f))
    finally:
        shutil.rmtree(c_dir, ignore_errors=True)


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


# 南兴优化清单 CSV 列（套板南兴优化清单.csv，GBK 编码；分柜号/特殊工艺为后期新增可选列）
CSV_COLS = ['条形码', '材料', '开料长', '开料宽', '数量', '厚度', '工件名称', '颜色',
            '订单编号', '客户地址', '批次号', '材料描述2', '分柜号', '特殊工艺']
# 管线标准输入列（门套/窗套/护墙 sheet，18 列）
XLS_HEADER = ['工件名称', '大类', '开料长', '开料宽', '数量', '厚度', '单双', '颜色',
              '生产号', '客户名称', '工艺', '材质', '内填', '生产类型', '单项备注',
              '特殊要求', '回传单号', '安装位置']
# 透传列：原样保留到最终分类输出（追加在 18 标准列之后）
PASSTHROUGH_COLS = ['条形码', '批次号', '材料描述2']
# 内部透传列：不追加到输出，仅用于覆盖输出「工件名称」列为原始整串
ORIG_ITEM_COL = '原工件名称'
# 新增透传列（PTB260917002 起 CSV 新增）：分柜号、特殊工艺，原样追加到输出 M/N 列
PASSTHROUGH_COLS2 = ['分柜号', '特殊工艺']


def _csv_sheet_kind(item_name):
    """根据 CSV 工件名称判断归入哪个 sheet：窗套/护墙/门套（默认）"""
    if '护墙' in item_name:
        return '护墙'
    if '窗套' in item_name or '哑口' in item_name or item_name.startswith('指接板'):
        return '窗套'
    return '门套'


def csv_to_xls(src_csv, dst_xls):
    """南兴优化清单 CSV 转管线标准 xls（门套/窗套/护墙 三 sheet 结构，值为空也保留表头）"""
    raw = open(src_csv, 'rb').read()
    for enc in ('utf-8-sig', 'gbk'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f'无法识别 CSV 编码: {src_csv}')

    rows = list(csv.reader(text.splitlines()))
    if not rows:
        raise ValueError(f'CSV 无内容: {src_csv}')
    header = [c.strip() for c in rows[0]]
    col_idx = {name: header.index(name) for name in CSV_COLS if name in header}
    for required in ('开料长', '开料宽', '数量', '厚度', '工件名称', '颜色'):
        if required not in col_idx:
            raise ValueError(f'CSV 缺少必要列「{required}」: {src_csv}')

    def get(row, name):
        i = col_idx.get(name)
        return row[i].strip() if i is not None and i < len(row) else ''

    def num(row, name):
        v = get(row, name)
        try:
            return float(v)
        except ValueError:
            return v

    sheets = {'门套': [], '窗套': [], '护墙': []}
    for row in rows[1:]:
        if not any(c.strip() for c in row):
            continue
        item_name = get(row, '工件名称')
        # 工件名称形如 门套_竖套板/单面单开槽(合页面开槽)：大类取下划线后、斜杠前
        if '_' in item_name:
            item_type, rest = item_name.split('_', 1)
        else:
            item_type, rest = item_name, ''
        category = rest.split('/')[0].strip()
        single_double = '双面' if '双面' in item_name else ('单面' if '单面' in item_name else '')
        sheets[_csv_sheet_kind(item_name)].append([
            item_type, category, num(row, '开料长'), num(row, '开料宽'), num(row, '数量'),
            num(row, '厚度'), single_double, get(row, '颜色'), get(row, '订单编号'),
            get(row, '客户地址'), '', '', '', '正常生产', '', '', '', '',
            # 透传列：条形码 / 批次号 / 材料描述2 / 原工件名称(内部覆盖用) / 分柜号 / 特殊工艺
            get(row, '条形码'), get(row, '批次号'), get(row, '材料描述2'), item_name,
            get(row, '分柜号'), get(row, '特殊工艺'),
        ])

    wb = xlwt.Workbook()
    for sheet_name in ('门套', '窗套', '护墙'):
        data = sheets[sheet_name]
        if not data and sheet_name == '护墙':
            continue  # 护墙无数据则不建 sheet（管线对护墙是可选项）
        ws = wb.add_sheet(sheet_name)
        for c, h in enumerate(XLS_HEADER + PASSTHROUGH_COLS + [ORIG_ITEM_COL] + PASSTHROUGH_COLS2):
            ws.write(0, c, h)
        for r, row_data in enumerate(data, start=1):
            for c, v in enumerate(row_data):
                ws.write(r, c, v)
    wb.save(dst_xls)
    print(f'  [CSV] 已转换为 xls: 门套={len(sheets["门套"])}, 窗套={len(sheets["窗套"])}, 护墙={len(sheets["护墙"])}')


def copy_sheet_to_writer(book, writer, sheet_name, new_name=None, skip_red=False):
    """复制 sheet 到 writer，可选跳过红色行（封边条等豁免行始终保留）"""
    old_sheet = book.sheet_by_name(sheet_name)
    new_sheet = writer.add_sheet(new_name or sheet_name)
    out_row = 0
    blank_style = xlwt.XFStyle()
    kept = 0
    for row in range(old_sheet.nrows):
        if skip_red and row > 0 and is_red_row(book, old_sheet, row):
            values = read_row_values(old_sheet, row)
            if is_keep_row(values):
                kept += 1
                add_kept(sheet_name, values)
            else:
                add_excluded(sheet_name, values, '红字标记')
                continue
        for col in range(old_sheet.ncols):
            value = old_sheet.cell_value(row, col)
            new_sheet.write(out_row, col, value, blank_style)
        out_row += 1
    if kept:
        print(f'  [Keep] {sheet_name} 标红但属豁免行(封边条等)，已保留: {kept} 行')
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

    # 统计红色行（封边条等豁免行会保留，不计入被过滤数）
    def _count_red(sheet):
        n = 0
        for r in range(1, sheet.nrows):
            if is_red_row(book_src, sheet, r) and not is_keep_row(read_row_values(sheet, r)):
                n += 1
        return n

    mk_red = _count_red(mk_sheet_src)
    ct_red = _count_red(ct_sheet_src)
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

    reset_events()  # 清空本线程的剔除/保留记录（回传到飞书卡片用）
    work_dir = Path(tempfile.mkdtemp(prefix='pvc_'))
    try:
        final_output = work_dir / f'{date_code}_for_classify.xls'
        print(f"{'='*60}")
        print(f"处理文件: {src_path}")
        print(f"工作目录: {work_dir}")
        print(f"{'='*60}")

        # 1. Pipeline（CSV 输入先转为标准 xls，输出需经宿主机 Excel 重写后南兴才能导入）
        is_csv = src_path.suffix.lower() == '.csv'
        if is_csv:
            converted = work_dir / f'{src_path.stem}_converted.xls'
            csv_to_xls(str(src_path), str(converted))
            run_pipeline(str(converted), str(final_output))
        else:
            run_pipeline(str(src_path), str(final_output))

        # 2. Classify
        auto_dir = work_dir / f'{dir_name}-自动分类'
        if auto_dir.exists():
            shutil.rmtree(auto_dir)
        auto_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n  [Classify] Running classify_output...")
        summary = classify_output.process_file(str(final_output), str(auto_dir), str(REF_DIR), str(COLOR_MAP)) or {}

        # CSV 来源的输出必须经宿主机 Excel 重写（南兴只认 Excel/WPS 重写过的文件）
        if is_csv:
            resave_via_host(str(auto_dir))

        output_files = sorted([f for f in auto_dir.rglob('*.xls') if f.is_file()])
        events = get_events()
        result = {
            'work_dir': str(work_dir),
            'auto_dir': str(auto_dir),
            'output_files': [str(f) for f in output_files],
            'count': len(output_files),
            'quantity_total': summary.get('quantity_total', 0),
            'quantity_files': summary.get('quantity_files', []),
            # 被剔除 / 被豁免保留的行：随响应回传，由网关写进飞书卡片
            'excluded_rows': events['excluded'],
            'kept_rows': events['kept'],
            'warnings': format_events(),
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
