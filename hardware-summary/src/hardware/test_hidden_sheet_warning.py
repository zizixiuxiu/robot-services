#!/usr/bin/env python3
"""Only warn when a hidden hardware sheet has real items."""
import os
import tempfile
import unittest

import xlwt

from convert import detect_hidden_source_sheets, format_hidden_sheet_warning

WU = '\u4e94'  # ��


def _add_wu_sheet(wb, name, hidden=False, item=None):
    ws = wb.add_sheet(name)
    ws.visibility = 1 if hidden else 0
    ws.write(0, 1, 'title')
    ws.write(1, 1, 'brand')
    ws.write(4, 1, 'name')
    if item:
        ws.write(5, 1, item['name'])
        ws.write(5, 7, item['qty'])
        ws.write(5, 9, item.get('unit', 'pcs'))
        ws.write(5, 11, '')
        ws.write(5, 14, '')
        ws.write(5, 18, '')
        ws.write(5, 19, '')
    return ws


class HiddenSheetWarningTest(unittest.TestCase):
    def test_visible_only_no_warning(self):
        wb = xlwt.Workbook()
        _add_wu_sheet(wb, '1' + WU, item={'name': 'hinge', 'qty': 2})
        path = os.path.join(tempfile.gettempdir(), 'hidden_warn_none.xls')
        wb.save(path)
        included = detect_hidden_source_sheets(path)
        self.assertEqual(included, [])
        self.assertIsNone(format_hidden_sheet_warning(included))

    def test_hidden_with_items_warns(self):
        wb = xlwt.Workbook()
        _add_wu_sheet(wb, '1' + WU, item={'name': 'hinge', 'qty': 2})
        _add_wu_sheet(wb, '2' + WU, hidden=True, item={'name': 'handle', 'qty': 4})
        path = os.path.join(tempfile.gettempdir(), 'hidden_warn_included.xls')
        wb.save(path)
        included = detect_hidden_source_sheets(path)
        self.assertEqual(included, ['2' + WU])
        msg = format_hidden_sheet_warning(included)
        self.assertIn('2' + WU, msg)
        self.assertIsNotNone(msg)

    def test_hidden_empty_no_warning(self):
        wb = xlwt.Workbook()
        _add_wu_sheet(wb, '1' + WU, item={'name': 'hinge', 'qty': 2})
        _add_wu_sheet(wb, '11' + WU + '\u91d1', hidden=True)
        path = os.path.join(tempfile.gettempdir(), 'hidden_warn_skipped.xls')
        wb.save(path)
        included = detect_hidden_source_sheets(path)
        self.assertEqual(included, [])
        self.assertIsNone(format_hidden_sheet_warning(included))

    def test_mixed_only_content_warns(self):
        wb = xlwt.Workbook()
        _add_wu_sheet(wb, '1' + WU, item={'name': 'hinge', 'qty': 2})
        _add_wu_sheet(wb, '2' + WU, hidden=True, item={'name': 'handle', 'qty': 4})
        _add_wu_sheet(wb, '11' + WU + '\u91d1', hidden=True)
        path = os.path.join(tempfile.gettempdir(), 'hidden_warn_mixed.xls')
        wb.save(path)
        included = detect_hidden_source_sheets(path)
        self.assertEqual(included, ['2' + WU])
        msg = format_hidden_sheet_warning(included)
        self.assertIn('2' + WU, msg)
        self.assertNotIn('11' + WU, msg)


if __name__ == '__main__':
    unittest.main()
