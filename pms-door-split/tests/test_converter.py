from __future__ import annotations

import csv
import io
import unittest

from app.converter import ConversionError, transform_rows, write_csv


HEADERS = [
    "订单号", "板件名称", "加工长度", "加工宽度", "数量", "材料描述", "纹理",
    "完工长度", "完工宽度", "打孔工艺", "特殊工艺", "正面条码", "开槽工艺",
    "客户", "柜体名称", "分柜号", "品牌", "厚度", "面积", "批次号", "工艺路线",
    "材料描述2", "物料编码", "异型", "订单类型", "封边类型",
]


def make_row(workpiece, quantity, desc, thickness, route, material_code="EN26-08-11-1054"):
    row = [""] * len(HEADERS)
    row[HEADERS.index("订单号")] = "PC001-2608-01P01-01"
    row[HEADERS.index("板件名称")] = workpiece
    row[HEADERS.index("数量")] = quantity
    row[HEADERS.index("材料描述")] = desc
    row[HEADERS.index("厚度")] = thickness
    row[HEADERS.index("工艺路线")] = route
    row[HEADERS.index("物料编码")] = material_code
    return row


class PmsConverterTests(unittest.TestCase):
    def test_double_thickness_row_is_split_and_quantity_doubled(self) -> None:
        rows = [make_row("YM-108", 1, "YSM0250451", 50, "复合贴YSM皮4+4MM")]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(output), 2)
        self.assertEqual([r[HEADERS.index("数量")] for r in output], [2, 2])
        self.assertEqual([r[HEADERS.index("材料描述")] for r in output], ["4mm素板", "4mm素板"])
        self.assertEqual([r[HEADERS.index("厚度")] for r in output], [4, 4])
        self.assertEqual(stats.split_source_rows, 1)
        self.assertEqual(stats.split_details[0], "订单 EN26-08-11-1054：4+4MM → 4mm, 4mm")

    def test_board_material_fuzzy_match(self) -> None:
        rows = [
            make_row("平板无造型", 1, "YSM23210", 50, "复合贴YSM皮8MM多层加密"),
            make_row("平板无造型", 1, "YSM23210", 50, "PVC系列28MM多层板加密"),
            make_row("平板无造型", 1, "YSM23210", 50, "黑碳晶贴YSM皮8MM"),
        ]
        output, stats = transform_rows(HEADERS, rows)

        descs = [r[HEADERS.index("材料描述")] for r in output]
        self.assertEqual(descs, ["8mm多层加密", "28mm多层加密", "8mm黑碳晶"])
        self.assertEqual(stats.material_mismatch_details, ())

    def test_excluded_model_and_keyword_rows_are_dropped(self) -> None:
        rows = [
            make_row("N1054(N1054)单扇", 1, "YSM2305", 50, "复合贴YSM皮8MM"),
            make_row("铝框门（灰框）", 1, "YSM2305", 50, "复合贴YSM皮8MM"),
            make_row("YM-035", 1, "YSM2301", 50, "复合贴YSM皮8MM"),
        ]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(output), 1)
        self.assertEqual(output[0][HEADERS.index("板件名称")], "YM-035")
        self.assertEqual(len(stats.excluded_details), 2)

    def test_note_row_without_workpiece_or_route_is_skipped(self) -> None:
        note = [""] * len(HEADERS)
        note[HEADERS.index("数量")] = "翻倍"
        note[HEADERS.index("材料描述")] = "多层加密 素板 碳晶板"
        rows = [note, make_row("YM-035", 1, "YSM2301", 50, "复合贴YSM皮8MM")]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(output), 1)
        self.assertEqual(stats.source_rows, 1)

    def test_blank_thickness_keeps_original_description(self) -> None:
        rows = [make_row("YM-035", 1, "YSM2303", 50, "复合贴YSM皮装板")]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(output[0][HEADERS.index("材料描述")], "YSM2303")
        self.assertEqual(output[0][HEADERS.index("厚度")], 50)
        self.assertEqual(len(stats.blank_thickness_orders), 1)

    def test_unseen_process_is_reported(self) -> None:
        rows = [make_row("YM-035", 1, "YSM2301", 50, "PET高光板8MM")]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(stats.new_process_details), 1)
        self.assertIn("PET高光板", stats.new_process_details[0])
        self.assertEqual(output[0][HEADERS.index("材料描述")], "8mm素板")

    def test_missing_header_is_rejected(self) -> None:
        bad_headers = [h for h in HEADERS if h != "工艺路线"]
        with self.assertRaisesRegex(ConversionError, "源表缺少字段"):
            transform_rows(bad_headers, [])

    def test_already_split_skin_row_passes_through_unchanged(self) -> None:
        rows = [make_row("YM-108", 2, "4mm素板", 4, "复合贴YSM皮4+4MM")]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(output), 1)
        self.assertEqual(output[0][HEADERS.index("数量")], 2)
        self.assertEqual(output[0][HEADERS.index("材料描述")], "4mm素板")
        self.assertEqual(stats.split_source_rows, 0)

    def test_csv_keeps_source_headers(self) -> None:
        data = write_csv(HEADERS, [make_row("YM-035", 2, "8mm素板", 8, "复合贴YSM皮8MM")])
        parsed = list(csv.reader(io.StringIO(data.decode("gb18030"))))
        self.assertEqual(parsed[0], HEADERS)
        self.assertEqual(parsed[1][HEADERS.index("材料描述")], "8mm素板")

    def test_csv_input_is_supported(self) -> None:
        from app.converter import convert_excel_to_csv

        buf = io.StringIO(newline="")
        writer = csv.writer(buf)
        writer.writerow(HEADERS)
        writer.writerow(make_row("YM-108", 1, "YSM0250451", 50, "复合贴YSM皮4+4MM"))
        content = buf.getvalue().encode("gb18030")

        csv_bytes, stats = convert_excel_to_csv(content, "清单.csv")
        rows = list(csv.reader(io.StringIO(csv_bytes.decode("gb18030"))))

        self.assertEqual(stats.source_rows, 1)
        self.assertEqual(len(rows) - 1, 2)
        self.assertEqual(rows[1][HEADERS.index("材料描述")], "4mm素板")
        self.assertEqual(rows[1][HEADERS.index("数量")], "2")


if __name__ == "__main__":
    unittest.main()
