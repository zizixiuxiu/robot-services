#!/usr/bin/env python3
"""Validate factory-version .xls output against the original hardware workbook.

The factory generator intentionally changes layout, hidden rows/columns, print
settings, and date cells/formula cells. This validator treats those known changes
as allowed and compares all other cell values strictly.
"""

import argparse
import json
import os
import re
import sys
from collections import defaultdict

import xlrd
from xlwt import Utils


DATE_RE = re.compile(r'^\d{1,4}[\.\-]\d{1,2}([\.\-]\d{1,2})?$')


def is_wu_sheet(sheet_name):
    return '五' in sheet_name or '五金' in sheet_name


def is_number_sheet(sheet_name):
    return sheet_name in [str(i) for i in range(1, 100)]


def cell_name(row, col):
    return Utils.rowcol_to_cell(row, col)


def row_text(sheet, row):
    vals = []
    for col in range(sheet.ncols):
        value = sheet.cell(row, col).value
        if value not in ('', None):
            vals.append(str(value))
    return ' '.join(vals)


def find_row_containing(sheet, keyword, start_row=0):
    for row in range(start_row, sheet.nrows):
        if keyword in row_text(sheet, row):
            return row
    return None


def is_date_value(book, sheet, row, col):
    cell_val = sheet.cell(row, col).value
    if not cell_val and cell_val != 0:
        return False
    if isinstance(cell_val, str) and DATE_RE.match(cell_val):
        return True
    if isinstance(cell_val, (int, float)) and 0 < cell_val < 100000:
        try:
            if sheet.cell(row, col).ctype == xlrd.XL_CELL_DATE:
                return True
        except Exception:
            pass
        xf_index = sheet.cell_xf_index(row, col)
        xf = book.xf_list[xf_index]
        fmt = book.format_map.get(xf.format_key) if hasattr(book, 'format_map') else None
        if fmt:
            fmt_str = fmt.format_str.lower()
            if 'date' in fmt_str or 'yy' in fmt_str or 'mm' in fmt_str or 'dd' in fmt_str:
                return True
    return False


def find_sheet1_date_ref(book, sheet, row, col):
    first_blank_col = None
    for dc in range(1, 6):
        next_col = col + dc
        if next_col >= sheet.ncols:
            break
        next_val = sheet.cell(row, next_col).value
        if is_date_value(book, sheet, row, next_col):
            return row, next_col
        if next_val:
            return row, next_col
        if first_blank_col is None:
            first_blank_col = next_col
    if first_blank_col is not None:
        return row, first_blank_col
    return None


def scan_sheet1_date_refs(book):
    refs = {}
    if '1' not in book.sheet_names():
        return refs
    sheet = book.sheet_by_name('1')
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            value = sheet.cell(row, col).value
            if not value or not isinstance(value, str):
                continue
            if '接单日期' in value and 'receipt_date' not in refs:
                ref = find_sheet1_date_ref(book, sheet, row, col)
                if ref:
                    refs['receipt_date'] = ref
            elif '下单日期' in value and 'order_date' not in refs:
                ref = find_sheet1_date_ref(book, sheet, row, col)
                if ref:
                    refs['order_date'] = ref
            elif (
                ('预计交货日期' in value or '预计交货期' in value or '包装预计交货' in value)
                and 'delivery_date' not in refs
            ):
                ref = find_sheet1_date_ref(book, sheet, row, col)
                if ref:
                    refs['delivery_date'] = ref
    return refs


def find_date_target_col(sheet, row, col, wu_sheet):
    if wu_sheet:
        for rlo, rhi, clo, chi in getattr(sheet, 'merged_cells', []):
            if rlo <= row < rhi and clo <= col < chi:
                if chi < sheet.ncols:
                    return chi
                return None

    target_col = None
    max_offset = 10 if wu_sheet else 6
    for dc in range(1, max_offset):
        next_col = col + dc
        if next_col >= sheet.ncols:
            break
        next_val = sheet.cell(row, next_col).value
        if next_val:
            return next_col
        if target_col is None:
            target_col = next_col
    return target_col


def expected_changed_cells(book):
    """Return allowed changed cells and expected formula targets.

    Returns:
        allowed_value_changes: set[(sheet_name, row, col)]
        expected_formulas: dict[(sheet_name, row, col)] -> formula_text
    """
    date_refs = scan_sheet1_date_refs(book)
    allowed_value_changes = set()
    expected_formulas = {}

    for sheet_name in book.sheet_names():
        sheet = book.sheet_by_name(sheet_name)
        wu_sheet = is_wu_sheet(sheet_name)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                value = sheet.cell(row, col).value
                if not value or not isinstance(value, str):
                    continue

                ref_key = None
                is_delivery = False
                if wu_sheet:
                    if '接单日期' in value:
                        ref_key = 'receipt_date'
                    elif '下单日期' in value:
                        ref_key = 'order_date'
                    elif '包装预计交货日期' in value:
                        is_delivery = True
                else:
                    if '接单日期' in value:
                        ref_key = 'receipt_date'
                    elif '下单日期' in value:
                        ref_key = 'order_date'
                    elif '包装预计交货期' in value:
                        is_delivery = True

                if ref_key is None and not is_delivery:
                    continue

                target_col = find_date_target_col(sheet, row, col, wu_sheet)
                if target_col is None:
                    continue

                key = (sheet_name, row, target_col)
                if sheet_name == '1':
                    allowed_value_changes.add(key)
                elif ref_key and ref_key in date_refs:
                    ref_row, ref_col = date_refs[ref_key]
                    expected_formulas[key] = "'1'!" + cell_name(ref_row, ref_col)
                    allowed_value_changes.add(key)
                elif ref_key and 'order_date' in date_refs:
                    ref_row, ref_col = date_refs['order_date']
                    expected_formulas[key] = "'1'!" + cell_name(ref_row, ref_col)
                    allowed_value_changes.add(key)
                elif is_delivery and 'delivery_date' in date_refs:
                    ref_row, ref_col = date_refs['delivery_date']
                    expected_formulas[key] = "'1'!" + cell_name(ref_row, ref_col)
                    allowed_value_changes.add(key)

    return allowed_value_changes, expected_formulas


def values_equal(left, right):
    if left == right:
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return abs(float(left) - float(right)) < 1e-9
    return False


def hidden_cols_for_sheet(sheet_name):
    if is_number_sheet(sheet_name):
        return {7, 8, 9, 10}
    if is_wu_sheet(sheet_name):
        return {10, 16, 17}
    if sheet_name == '特殊产品表' or sheet_name.startswith('外购表'):
        return {10, 11, 12}
    if sheet_name == '实木':
        return {10, 11}
    if sheet_name == '汇总':
        return {1, 2, 3, 4}
    return set()


def should_hide_row_digital(sheet, row_idx):
    text = row_text(sheet, row_idx)
    if '板材' in text or '五金' in text:
        return True
    if '此单共' in text or '合计金额' in text:
        return True
    if '经销商确认' in text:
        return True
    return False


def should_hide_row_hardware(sheet, row_idx):
    text = row_text(sheet, row_idx)
    if '此单共' in text or '合计金额' in text:
        return True
    if '优惠' in text or '需支付' in text:
        return True
    return False


def is_process_area_row(sheet, row_idx):
    text = row_text(sheet, row_idx)
    return '柜体加工' in text or '门板加工' in text


def is_zhibiaoren_row(sheet, row_idx):
    text = row_text(sheet, row_idx)
    if '制表人' not in text:
        return False
    for rr in range(max(0, row_idx - 8), row_idx):
        if is_process_area_row(sheet, rr):
            return True
    return False


def is_blank_tuwen_box(sheet, row_idx):
    text = row_text(sheet, row_idx)
    if '图文说明:' not in text:
        return False
    for rr in range(row_idx + 1, min(sheet.nrows, row_idx + 10)):
        if is_process_area_row(sheet, rr):
            return True
    return False


def normal_table_rows(sheet):
    start = find_row_containing(sheet, '序号')
    if start is None:
        return []
    stop_keywords = (
        '图文说明', '定金', '此单共', '合计金额', '经销商确认', '制表人',
        '柜体加工', '门板加工', '下单日期', '包装预计交货'
    )
    end = sheet.nrows - 1
    for row in range(start + 1, sheet.nrows):
        text = row_text(sheet, row)
        if any(keyword in text for keyword in stop_keywords):
            end = row - 1
            break
    return list(range(start, end + 1))


def validate_workbook(original_path, generated_path, check_layout=True, max_examples=30):
    original = xlrd.open_workbook(original_path, formatting_info=True)
    generated = xlrd.open_workbook(generated_path, formatting_info=True)
    allowed_changes, expected_formulas = expected_changed_cells(original)

    errors = []
    warnings = []
    stats = defaultdict(int)

    if original.sheet_names() != generated.sheet_names():
        errors.append({
            'kind': 'sheet_names_mismatch',
            'original': original.sheet_names(),
            'generated': generated.sheet_names(),
        })
        return errors, warnings, stats

    for sheet_name in original.sheet_names():
        src = original.sheet_by_name(sheet_name)
        out = generated.sheet_by_name(sheet_name)
        stats['sheets_checked'] += 1

        if src.nrows != out.nrows or src.ncols != out.ncols:
            errors.append({
                'kind': 'shape_mismatch',
                'sheet': sheet_name,
                'original': [src.nrows, src.ncols],
                'generated': [out.nrows, out.ncols],
            })
            continue

        for row in range(src.nrows):
            for col in range(src.ncols):
                key = (sheet_name, row, col)
                if key in allowed_changes:
                    stats['allowed_changed_cells'] += 1
                    continue
                stats['cells_compared'] += 1
                left = src.cell(row, col).value
                right = out.cell(row, col).value
                if not values_equal(left, right):
                    if len(errors) < max_examples:
                        errors.append({
                            'kind': 'value_mismatch',
                            'sheet': sheet_name,
                            'cell': cell_name(row, col),
                            'original': left,
                            'generated': right,
                        })
                    stats['value_mismatches'] += 1

        for key, formula_text in expected_formulas.items():
            formula_sheet, row, col = key
            if formula_sheet != sheet_name:
                continue
            stats['expected_formula_cells'] += 1
            # xlrd exposes xlwt-written formula cells as cached values, often ''.
            # We cannot reliably read formula text back from xlrd, so this check
            # verifies the expected target location is allowed to differ and was
            # not overwritten with the old source value.
            src_val = src.cell(row, col).value
            out_val = out.cell(row, col).value
            if src_val not in ('', None) and values_equal(src_val, out_val):
                warnings.append({
                    'kind': 'formula_cell_still_has_original_cached_value',
                    'sheet': sheet_name,
                    'cell': cell_name(row, col),
                    'expected_formula': formula_text,
                    'value': out_val,
                })

        if check_layout:
            for col in hidden_cols_for_sheet(sheet_name):
                if col >= out.ncols:
                    continue
                colinfo = out.colinfo_map.get(col)
                if not colinfo or not colinfo.hidden:
                    errors.append({
                        'kind': 'expected_hidden_col_not_hidden',
                        'sheet': sheet_name,
                        'col': cell_name(0, col).rstrip('1'),
                    })

            if is_wu_sheet(sheet_name):
                for col, expected_width in [(6, 15 * 256), (9, 11 * 256)]:
                    if col < out.ncols:
                        colinfo = out.colinfo_map.get(col)
                        if not colinfo or colinfo.width != expected_width:
                            errors.append({
                                'kind': 'wu_col_width_mismatch',
                                'sheet': sheet_name,
                                'col': cell_name(0, col).rstrip('1'),
                                'expected': expected_width,
                                'actual': colinfo.width if colinfo else None,
                            })
                for row in range(out.nrows):
                    if should_hide_row_hardware(src, row):
                        rowinfo = out.rowinfo_map.get(row)
                        if not rowinfo or not rowinfo.hidden or rowinfo.height != 0:
                            errors.append({
                                'kind': 'expected_hidden_wu_row_not_hidden',
                                'sheet': sheet_name,
                                'row': row + 1,
                                'text': row_text(src, row),
                            })

            if is_number_sheet(sheet_name):
                for row in normal_table_rows(src):
                    rowinfo = out.rowinfo_map.get(row)
                    if not rowinfo or rowinfo.height != 30 * 20:
                        errors.append({
                            'kind': 'normal_table_row_height_mismatch',
                            'sheet': sheet_name,
                            'row': row + 1,
                            'expected': 30 * 20,
                            'actual': rowinfo.height if rowinfo else None,
                        })
                for row in range(out.nrows):
                    if is_process_area_row(src, row) or is_zhibiaoren_row(src, row) or is_blank_tuwen_box(src, row):
                        continue
                    if should_hide_row_digital(src, row):
                        rowinfo = out.rowinfo_map.get(row)
                        if not rowinfo or not rowinfo.hidden or rowinfo.height != 0:
                            errors.append({
                                'kind': 'expected_hidden_normal_row_not_hidden',
                                'sheet': sheet_name,
                                'row': row + 1,
                                'text': row_text(src, row),
                            })

    return errors, warnings, stats


def main():
    parser = argparse.ArgumentParser(description='Validate factory .xls output against original workbook.')
    parser.add_argument('original')
    parser.add_argument('generated')
    parser.add_argument('--no-layout', action='store_true', help='Only validate cell values, skip layout/hidden checks.')
    parser.add_argument('--json', action='store_true', help='Print machine-readable JSON.')
    parser.add_argument('--max-examples', type=int, default=30)
    args = parser.parse_args()

    errors, warnings, stats = validate_workbook(
        args.original,
        args.generated,
        check_layout=not args.no_layout,
        max_examples=args.max_examples,
    )

    result = {
        'ok': not errors,
        'original': os.path.abspath(args.original),
        'generated': os.path.abspath(args.generated),
        'stats': dict(stats),
        'errors': errors,
        'warnings': warnings[:args.max_examples],
        'warning_count': len(warnings),
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print('OK' if result['ok'] else 'FAILED')
        print('original:', result['original'])
        print('generated:', result['generated'])
        print('stats:', json.dumps(result['stats'], ensure_ascii=False))
        if warnings:
            print('warnings:', len(warnings))
            for warning in warnings[:args.max_examples]:
                print('  WARN', warning)
        if errors:
            print('errors:', len(errors))
            for error in errors[:args.max_examples]:
                print('  ERR', error)

    return 0 if result['ok'] else 1


if __name__ == '__main__':
    sys.exit(main())
