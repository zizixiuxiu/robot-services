#!/usr/bin/env python3
"""Do not double-count vertically merged hardware rows; checksum vs 数量总计."""
import os
import tempfile
import unittest

import xlrd
import xlwt

from convert import (
    _extract_data_from_hardware_sheet,
    check_hardware_quantity_totals,
    extract_hardware_data,
    format_quantity_mismatch_warning,
)

WU = '\u4e94'


def _pad_row(ws, row, unit='pcs'):
    ws.write(row, 9, unit)
    ws.write(row, 11, '')
    ws.write(row, 14, '')
    ws.write(row, 18, '')
    ws.write(row, 19, '')


def _write_qty_total(ws, row, total):
    ws.write_merge(row, row, 3, 6, '数量总计：')
    ws.write_merge(row, row, 7, 8, total)


class MergedNameRowTest(unittest.TestCase):
    def test_vertical_name_merge_counts_once(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('6' + WU)
        ws.write(4, 1, 'name')
        ws.write(5, 1, 'jiaojiaoma')
        ws.write(5, 7, 16)
        _pad_row(ws, 5)
        ws.write_merge(6, 7, 1, 6, '3- WJXZ-01 screw 4*16')
        ws.write_merge(6, 7, 7, 8, 64)
        _pad_row(ws, 6)
        _pad_row(ws, 7)
        _write_qty_total(ws, 9, 80)
        path = os.path.join(tempfile.gettempdir(), 'hw_merged_name.xls')
        wb.save(path)

        items = extract_hardware_data(path)
        self.assertEqual(len(items), 2)
        self.assertEqual(sum(i['quantity'] for i in items), 80)
        names = [i['hardware_name'] for i in items]
        self.assertEqual(names.count('3- WJXZ-01 screw 4*16'), 1)
        self.assertEqual(check_hardware_quantity_totals(path), [])

    def test_vertical_qty_merge_with_repeated_name_counts_once(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('6' + WU)
        ws.write(4, 1, 'name')
        ws.write(5, 1, 'screw')
        ws.write(6, 1, 'screw')
        ws.write_merge(5, 6, 7, 8, 64)
        _pad_row(ws, 5)
        _pad_row(ws, 6)
        _write_qty_total(ws, 8, 64)
        path = os.path.join(tempfile.gettempdir(), 'hw_merged_qty.xls')
        wb.save(path)

        items = extract_hardware_data(path)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]['quantity'], 64)
        self.assertEqual(check_hardware_quantity_totals(path), [])

    def test_horizontal_name_merge_same_row_still_counts(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('1' + WU)
        ws.write(4, 1, 'name')
        ws.write_merge(5, 5, 1, 6, 'hinge')
        ws.write(5, 7, 2)
        _pad_row(ws, 5)
        path = os.path.join(tempfile.gettempdir(), 'hw_horiz_name.xls')
        wb.save(path)

        sheet = xlrd.open_workbook(path, formatting_info=True).sheet_by_index(0)
        items = _extract_data_from_hardware_sheet(sheet)
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]['quantity'], 2)

    def test_checksum_warns_when_extracted_differs_from_sheet_total(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('6' + WU)
        ws.write(4, 1, 'name')
        ws.write(5, 1, 'jiaojiaoma')
        ws.write(5, 7, 16)
        _pad_row(ws, 5)
        ws.write(6, 1, 'screw')
        ws.write(6, 7, 64)
        _pad_row(ws, 6)
        _write_qty_total(ws, 8, 999)
        path = os.path.join(tempfile.gettempdir(), 'hw_qty_mismatch.xls')
        wb.save(path)

        mismatches = check_hardware_quantity_totals(path)
        self.assertEqual(len(mismatches), 1)
        self.assertEqual(mismatches[0]['sheet'], '6' + WU)
        self.assertEqual(mismatches[0]['extracted'], 80)
        self.assertEqual(mismatches[0]['labeled'], 999)
        msg = format_quantity_mismatch_warning(mismatches)
        self.assertIn('数量核对失败', msg)
        self.assertIn('抽出80', msg)
        self.assertIn('表内999', msg)


if __name__ == '__main__':
    unittest.main()
