#!/usr/bin/env python3
"""1五 隐藏时，日期公式锚点应落到第一个可见五金表。"""
import os
import tempfile
import unittest

import xlrd
import xlwt

from hide_prices import _get_date_source_sheet_name, generate_factory_version


def _add_normal_sheet(wb, name, hidden=False):
    ws = wb.add_sheet(name)
    ws.visibility = 1 if hidden else 0
    ws.write(0, 0, '下单日期')
    ws.write(0, 1, '2026.8.1')
    ws.write(1, 0, '包装预计交货期')
    ws.write(1, 1, '2026.8.10')
    return ws


def _add_wu_sheet(wb, name, hidden=False):
    ws = wb.add_sheet(name)
    ws.visibility = 1 if hidden else 0
    ws.write(0, 0, '下单日期')
    ws.write(0, 1, '2026.8.1')
    ws.write(1, 0, '包装预计交货日期')
    ws.write(1, 1, '2026.8.10')
    return ws


class HiddenDateSourceTest(unittest.TestCase):
    def test_visible_wu_still_preferred(self):
        wb = xlwt.Workbook()
        _add_normal_sheet(wb, '1')
        _add_wu_sheet(wb, '1五')
        path = os.path.join(tempfile.gettempdir(), 'date_src_visible_wu.xls')
        wb.save(path)
        book = xlrd.open_workbook(path, formatting_info=True)
        self.assertEqual(_get_date_source_sheet_name(book), '1五')

    def test_hidden_wu_falls_back_to_first_visible_wu(self):
        wb = xlwt.Workbook()
        _add_normal_sheet(wb, '1')
        _add_wu_sheet(wb, '1五', hidden=True)
        _add_wu_sheet(wb, '2五')
        path = os.path.join(tempfile.gettempdir(), 'date_src_hidden_wu.xls')
        wb.save(path)
        book = xlrd.open_workbook(path, formatting_info=True)
        self.assertEqual(book.sheet_by_name('1五').visibility, 1)
        self.assertEqual(_get_date_source_sheet_name(book), '2五')

    def test_zhang_jianjun_source_file(self):
        candidates = [
            '/app/data/output/_src_zhang.xls',
            r'D:\Services\robot-services\hardware-summary\data\output\_src_zhang.xls',
            r'D:\飞书下载\B2607-1572张建军.xls',
        ]
        path = next((p for p in candidates if os.path.exists(p)), None)
        if not path:
            self.skipTest('张建军源文件不在')
        book = xlrd.open_workbook(path, formatting_info=True)
        self.assertEqual(book.sheet_by_name('1五').visibility, 1)
        self.assertEqual(book.sheet_by_name('2五').visibility, 0)
        self.assertEqual(_get_date_source_sheet_name(book), '2五')

    def test_factory_output_anchors_visible_wu_sheet(self):
        src = os.path.join(tempfile.gettempdir(), 'date_src_hidden_wu_in.xls')
        out = os.path.join(tempfile.gettempdir(), 'date_src_hidden_wu_out.xls')
        wb = xlwt.Workbook()
        _add_normal_sheet(wb, '1')
        _add_wu_sheet(wb, '1五', hidden=True)
        _add_wu_sheet(wb, '2五')
        wb.save(src)
        generate_factory_version(src, out, order_date='2026.8.18')
        book = xlrd.open_workbook(out, formatting_info=True)
        self.assertEqual(_get_date_source_sheet_name(book), '2五')
        sheet_wu = book.sheet_by_name('2五')
        self.assertEqual(str(sheet_wu.cell(0, 1).value).strip(), '2026.8.18')


if __name__ == '__main__':
    unittest.main()
