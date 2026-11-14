# -*- coding: utf-8 -*-
import os
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC_DIR))

from workshop_order_http import (  # noqa: E402
    archive_workshop_pair,
    canonical_archive_name,
    sanitize_uploader_name,
)


class ArchiveWorkshopTests(unittest.TestCase):
    def test_sanitize_uploader_name(self) -> None:
        self.assertEqual(sanitize_uploader_name('\u9648\u660c\u5bb9'), '\u9648\u660c\u5bb9')
        self.assertEqual(sanitize_uploader_name('  \u8c22\u7075  '), '\u8c22\u7075')
        self.assertEqual(sanitize_uploader_name('a/b:c*d'), 'abcd')
        self.assertEqual(sanitize_uploader_name(''), '')
        self.assertEqual(sanitize_uploader_name('\u672a\u77e5\u4e0a\u4f20\u8005'), '')

    def test_canonical_archive_name_strips_windows_dup_suffix(self) -> None:
        self.assertEqual(
            canonical_archive_name('S2608-2529\u4e0b\u8f66\u95f4 (1).xlsx'),
            'S2608-2529\u4e0b\u8f66\u95f4.xlsx',
        )
        self.assertEqual(
            canonical_archive_name('S2607-1551\u8d34\u76ae(2).xls'),
            'S2607-1551\u8d34\u76ae.xls',
        )
        self.assertEqual(
            canonical_archive_name('S2608-2529\u6ed4\u6587\u65b0\uff08\u6df7\u6cb9\uff09.xls'),
            'S2608-2529\u6ed4\u6587\u65b0\uff08\u6df7\u6cb9\uff09.xls',
        )

    def test_archive_same_name_keeps_latest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'archive'
            first_in = Path(tmp) / 'old.xlsx'
            first_out = Path(tmp) / 'old\u4e0b\u8f66\u95f4.xlsx'
            first_in.write_bytes(b'OLD')
            first_out.write_bytes(b'OLD-OUT')
            archive_workshop_pair(
                first_in,
                'S2608-7520.xlsx',
                first_out,
                'S2608-7520\u4e0b\u8f66\u95f4.xlsx',
                '\u65b9\u5b87',
                now=datetime(2026, 9, 1, 9, 58),
                root=root,
            )
            later_in = Path(tmp) / 'new.xlsx'
            later_out = Path(tmp) / 'new\u4e0b\u8f66\u95f4.xlsx'
            later_in.write_bytes(b'NEW')
            later_out.write_bytes(b'NEW-OUT')
            archive_workshop_pair(
                later_in,
                'S2608-7520.xlsx',
                later_out,
                'S2608-7520\u4e0b\u8f66\u95f4.xlsx',
                '\u65b9\u5b87',
                now=datetime(2026, 9, 1, 10, 15),
                root=root,
            )
            src = root / '2026' / '9' / '\u65b9\u5b87' / '\u539f\u6587\u4ef6' / 'S2608-7520.xlsx'
            out = root / '2026' / '9' / '\u65b9\u5b87' / '\u5904\u7406\u540e' / 'S2608-7520\u4e0b\u8f66\u95f4.xlsx'
            self.assertEqual(src.read_bytes(), b'NEW')
            self.assertEqual(out.read_bytes(), b'NEW-OUT')
            self.assertEqual(len(list((root / '2026' / '9' / '\u65b9\u5b87' / '\u539f\u6587\u4ef6').iterdir())), 1)

    def test_archive_replaces_windows_dup_suffix_copy(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'archive'
            person = root / '2026' / '9' / '\u8c22\u7075' / '\u5904\u7406\u540e'
            person.mkdir(parents=True)
            old = person / 'S2608-2529\u4e0b\u8f66\u95f4.xlsx'
            dup = person / 'S2608-2529\u4e0b\u8f66\u95f4 (1).xlsx'
            old.write_bytes(b'OLD')
            dup.write_bytes(b'DUP')
            original = Path(tmp) / 'in.xls'
            processed = Path(tmp) / 'out.xlsx'
            original.write_bytes(b'SRC')
            processed.write_bytes(b'NEW')
            archive_workshop_pair(
                original,
                'S2608-2529.xls',
                processed,
                'S2608-2529\u4e0b\u8f66\u95f4.xlsx',
                '\u8c22\u7075',
                now=datetime(2026, 9, 1, 17, 0),
                root=root,
            )
            files = sorted(p.name for p in person.iterdir())
            self.assertEqual(files, ['S2608-2529\u4e0b\u8f66\u95f4.xlsx'])
            self.assertEqual((person / files[0]).read_bytes(), b'NEW')

    def test_archive_pair_creates_original_and_processed_dirs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'archive'
            original = Path(tmp) / 'S2608-4042.xlsx'
            processed = Path(tmp) / 'S2608-4042\u4e0b\u8f66\u95f4.xlsx'
            original.write_bytes(b'SRC')
            processed.write_bytes(b'OUT')
            person = archive_workshop_pair(
                original,
                'S2608-4042.xlsx',
                processed,
                'S2608-4042\u4e0b\u8f66\u95f4.xlsx',
                '\u9648\u660c\u5bb9',
                now=datetime(2026, 9, 1, 10, 0),
                root=root,
            )
            self.assertEqual(person, root / '2026' / '9' / '\u9648\u660c\u5bb9')
            src = root / '2026' / '9' / '\u9648\u660c\u5bb9' / '\u539f\u6587\u4ef6' / 'S2608-4042.xlsx'
            out = root / '2026' / '9' / '\u9648\u660c\u5bb9' / '\u5904\u7406\u540e' / 'S2608-4042\u4e0b\u8f66\u95f4.xlsx'
            self.assertEqual(src.read_bytes(), b'SRC')
            self.assertEqual(out.read_bytes(), b'OUT')

    def test_archive_skips_before_september_first(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'archive'
            original = Path(tmp) / 'a.xlsx'
            processed = Path(tmp) / 'a\u4e0b\u8f66\u95f4.xlsx'
            original.write_bytes(b'SRC')
            processed.write_bytes(b'OUT')
            person = archive_workshop_pair(
                original,
                'a.xlsx',
                processed,
                'a\u4e0b\u8f66\u95f4.xlsx',
                '\u9648\u660c\u5bb9',
                now=datetime(2026, 8, 31, 23, 0),
                root=root,
            )
            self.assertIsNone(person)
            self.assertFalse((root / '2026' / '9' / '\u9648\u660c\u5bb9').exists())

    def test_archive_skips_missing_uploader(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'archive'
            src = Path(tmp) / 'in.xlsx'
            out = Path(tmp) / 'out.xlsx'
            src.write_bytes(b'IN')
            out.write_bytes(b'OUT')
            self.assertIsNone(
                archive_workshop_pair(
                    src,
                    'x.xlsx',
                    out,
                    'x\u4e0b\u8f66\u95f4.xlsx',
                    '',
                    now=datetime(2026, 9, 2),
                    root=root,
                )
            )
            self.assertIsNone(
                archive_workshop_pair(
                    src,
                    'x.xlsx',
                    out,
                    'x\u4e0b\u8f66\u95f4.xlsx',
                    '\u672a\u77e5\u4e0a\u4f20\u8005',
                    now=datetime(2026, 9, 2),
                    root=root,
                )
            )
            self.assertFalse(root.exists())

    def test_archive_disabled_without_root(self) -> None:
        self.assertIsNone(
            archive_workshop_pair(
                Path('missing.xlsx'),
                'a.xlsx',
                Path('missing-out.xlsx'),
                'a\u4e0b\u8f66\u95f4.xlsx',
                '\u9648\u660c\u5bb9',
                now=datetime(2026, 9, 1),
                root=None,
            )
        )


    def test_archive_year_month_boundary_and_no_zero_padding(self) -> None:
        """跨年/跨月都按处理日期分层；年、月都不补零（2026/9、2027/1，不是 09/01）。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'archive'

            def archive(when: datetime, tag: str) -> Path | None:
                src = Path(tmp) / f'{tag}.xlsx'
                out = Path(tmp) / f'{tag}下车间.xlsx'
                src.write_bytes(b'SRC')
                out.write_bytes(b'OUT')
                return archive_workshop_pair(
                    src,
                    f'{tag}.xlsx',
                    out,
                    f'{tag}下车间.xlsx',
                    '方宇',
                    now=when,
                    root=root,
                )

            self.assertEqual(archive(datetime(2026, 9, 30, 23, 30), 'A'), root / '2026' / '9' / '方宇')
            self.assertEqual(archive(datetime(2026, 10, 1, 0, 30), 'B'), root / '2026' / '10' / '方宇')
            self.assertEqual(archive(datetime(2026, 12, 31, 23, 30), 'C'), root / '2026' / '12' / '方宇')
            self.assertEqual(archive(datetime(2027, 1, 1, 0, 30), 'D'), root / '2027' / '1' / '方宇')
            self.assertTrue((root / '2027' / '1' / '方宇' / '原文件' / 'D.xlsx').is_file())
            self.assertFalse((root / '09').exists())
            self.assertFalse((root / '01').exists())


if __name__ == '__main__':
    unittest.main()
