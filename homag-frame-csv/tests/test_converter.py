from __future__ import annotations

import csv
import io
import unittest

from app.color import convert as convert_color
from app.converter import (
    OUTPUT_HEADERS,
    ConversionError,
    convert_excel_to_csv,
    transform_rows,
    write_csv,
)


SOURCE_HEADERS = [
    "工件名称", "大类", "开料长", "开料宽", "数量", "厚度", "单双", "颜色",
    "生产号", "客户名称", "工艺", "材质", "内填",
]


def make_row(
    category="竖套板",
    length=2210,
    width=130,
    qty=2,
    thickness=28,
    face="双面 ",
    color="8-14G 平板/YSM232-6",
    customer="云南省.曲靖市.蒋东.",
    material="8MM",
) -> list:
    row = [""] * len(SOURCE_HEADERS)
    row[SOURCE_HEADERS.index("大类")] = category
    row[SOURCE_HEADERS.index("开料长")] = length
    row[SOURCE_HEADERS.index("开料宽")] = width
    row[SOURCE_HEADERS.index("数量")] = qty
    row[SOURCE_HEADERS.index("厚度")] = thickness
    row[SOURCE_HEADERS.index("单双")] = face
    row[SOURCE_HEADERS.index("颜色")] = color
    row[SOURCE_HEADERS.index("客户名称")] = customer
    row[SOURCE_HEADERS.index("材质")] = material
    return row


class ColorConvertTests(unittest.TestCase):
    def test_py_and_zky_and_ysm(self) -> None:
        self.assertEqual(convert_color("PY13")[0], "YSM230-5")
        self.assertEqual(convert_color("ZKY-5007")[0], "YSM230-1")
        self.assertEqual(convert_color("YSM-232-5")[0], "YSM232-5")
        self.assertEqual(convert_color("YSM224-17多层加密")[0], "YSM224-17多层加密")
        self.assertEqual(convert_color("MF37"), ("MF37", "MF37"))
        self.assertEqual(convert_color("琳琅白"), ("琳琅白", "琳琅白"))


class HomagConverterTests(unittest.TestCase):
    def test_row_mapping(self) -> None:
        rows = [make_row()]
        output, skipped, no_slash, qty = transform_rows(SOURCE_HEADERS, rows)
        self.assertEqual(skipped, 0)
        self.assertEqual(no_slash, 0)
        self.assertEqual(qty, 2)
        row = output[0]
        self.assertEqual(row[0], "8-14G 平板")
        self.assertEqual(row[1], "竖套板")
        self.assertEqual(row[2], 2210)
        self.assertEqual(row[3], 130)
        self.assertEqual(row[4], 2)
        self.assertEqual(row[5], "YSM232-6")
        self.assertEqual(row[6], 1)
        self.assertEqual(row[7], 2210)
        self.assertEqual(row[8], 130)
        self.assertEqual(row[10], "双面 8MM")
        self.assertEqual(row[13], "云南省.曲靖市.蒋东.")
        self.assertEqual(row[17], 28)
        self.assertEqual(row[21], "YSM232-6")

    def test_py_goes_to_material_and_original_stays(self) -> None:
        rows = [make_row(color="8-14G 造型/PY13")]
        output, _, _, _ = transform_rows(SOURCE_HEADERS, rows)
        self.assertEqual(output[0][5], "YSM230-5")
        self.assertEqual(output[0][21], "PY13")

    def test_empty_category_or_color_skipped(self) -> None:
        rows = [
            make_row(category=" ", color="8-14G 平板/YSM232-6"),
            make_row(category="竖套板", color=""),
            make_row(),
        ]
        output, skipped, _, _ = transform_rows(SOURCE_HEADERS, rows)
        self.assertEqual(len(output), 1)
        self.assertEqual(skipped, 2)

    def test_color_without_slash_keeps_empty_order(self) -> None:
        rows = [make_row(color="YSM224-15")]
        output, _, no_slash, _ = transform_rows(SOURCE_HEADERS, rows)
        self.assertEqual(no_slash, 1)
        self.assertEqual(output[0][0], "")
        self.assertEqual(output[0][5], "YSM224-15")
        self.assertEqual(output[0][21], "YSM224-15")

    def test_missing_header_is_rejected(self) -> None:
        bad = [h for h in SOURCE_HEADERS if h != "颜色"]
        with self.assertRaisesRegex(ConversionError, "源表缺少字段"):
            transform_rows(bad, [make_row()])

    def test_csv_roundtrip(self) -> None:
        data = write_csv(OUTPUT_HEADERS, transform_rows(SOURCE_HEADERS, [make_row()])[0])
        parsed = list(csv.reader(io.StringIO(data.decode("gb18030"))))
        self.assertEqual(parsed[0], OUTPUT_HEADERS)
        self.assertEqual(parsed[1][0], "8-14G 平板")
        self.assertEqual(parsed[1][5], "YSM232-6")

    def test_csv_input_is_supported(self) -> None:
        buf = io.StringIO(newline="")
        writer = csv.writer(buf)
        writer.writerow(SOURCE_HEADERS)
        writer.writerow(make_row(color="8-14H 平板/ZKY-5007"))
        writer.writerow(make_row(category="横套板", color="8-14H 平板/PY03"))
        content = buf.getvalue().encode("gb18030")

        csv_bytes, stats = convert_excel_to_csv(content, "汇总.csv")
        rows = list(csv.reader(io.StringIO(csv_bytes.decode("gb18030"))))
        self.assertEqual(stats.output_rows, 2)
        self.assertEqual(stats.door_rows, 2)
        self.assertEqual(rows[1][5], "YSM230-1")
        self.assertEqual(rows[1][21], "ZKY-5007")
        self.assertEqual(rows[2][5], "YSM230-1")
        self.assertEqual(rows[2][21], "PY03")
        self.assertFalse(stats.has_huqiang_file)

    def test_huqiang_copied_and_shoucai_dropped(self) -> None:
        from openpyxl import Workbook, load_workbook

        from app.converter import convert_workbook

        wb = Workbook()
        door = wb.active
        door.title = "门套"
        door.append(SOURCE_HEADERS)
        door.append(make_row(category="竖套板", color="8-14G 平板/PY13"))
        window = wb.create_sheet("窗套")
        window.append(SOURCE_HEADERS)
        window.append(make_row(category="横套板", color="8-14G 平板/YSM232-6"))
        wall = wb.create_sheet("护墙")
        wall.append(SOURCE_HEADERS)
        wall_row = make_row(category=" ", color="8-14G 平板/YSM230-5")
        wall_row[SOURCE_HEADERS.index("工件名称")] = "护墙"
        wall.append(wall_row)
        hand = wb.create_sheet("手裁")
        hand.append(SOURCE_HEADERS)
        hand.append(make_row(category="竖套板", color="8-14G 造型/MF37"))
        buf = io.BytesIO()
        wb.save(buf)

        result = convert_workbook(buf.getvalue(), "汇总.xlsx")
        csv_rows = list(csv.reader(io.StringIO(result.csv_bytes.decode("gb18030"))))
        self.assertEqual(result.stats.output_rows, 2)
        self.assertEqual(result.stats.huqiang_rows, 1)
        self.assertEqual(result.stats.discarded_shoucai, 1)
        self.assertTrue(result.stats.has_huqiang_file)
        self.assertIsNotNone(result.huqiang_bytes)
        self.assertEqual([r[1] for r in csv_rows[1:]], ["竖套板", "横套板"])
        self.assertNotIn("MF37", result.csv_bytes.decode("gb18030"))

        copied = load_workbook(io.BytesIO(result.huqiang_bytes))
        self.assertEqual(copied.sheetnames, ["护墙"])
        copied_rows = list(copied.active.iter_rows(values_only=True))
        self.assertEqual(copied_rows[0][SOURCE_HEADERS.index("颜色")], "颜色")
        self.assertEqual(copied_rows[1][SOURCE_HEADERS.index("颜色")], "8-14G 平板/YSM230-5")
        self.assertEqual(copied_rows[1][0], "护墙")


if __name__ == "__main__":
    unittest.main()
