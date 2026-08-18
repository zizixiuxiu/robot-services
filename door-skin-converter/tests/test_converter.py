from __future__ import annotations

import csv
import io
import unittest
from pathlib import Path

import app.converter as converter
from app.converter import ConversionError, OUTPUT_HEADERS, transform_rows, write_csv


HEADERS = [
    "材料",
    "开料长",
    "开料宽",
    "数量",
    "厚度",
    "工件名称",
    "颜色",
    "订单编号",
    "客户地址",
    "工艺",
    "特殊要求",
    "品牌",
]


class ConverterTests(unittest.TestCase):
    def setUp(self) -> None:
        # 隔离运行时剔除清单，避免宿主机 data 目录里的群命令结果影响用例
        converter.RUNTIME_EXCLUSIONS_FILE = Path("/tmp/__nonexistent_runtime_exclusions_test__.json")

    def test_composite_thickness_is_split_and_quantity_is_doubled(self) -> None:
        rows = [[
            "门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
            "ORDER-1", "重庆市.重庆市.客户", "5+8MM双贴面",
            "不开锁孔,加方钢（美心）,5+8内压线工艺", "美心蒙迪",
        ]]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(output), 2)
        self.assertEqual([row[4] for row in output], [2, 2])
        self.assertEqual([row[5] for row in output], [
            "5mm素板", "8mm素板"
        ])
        self.assertEqual([row[17] for row in output], [5, 8])
        self.assertEqual(
            output[0][10],
            "门扇厚度50mm.不开锁孔,加方钢（美心）,5+8内压线工艺",
        )
        self.assertEqual([row[21] for row in output], [
            "5mmPY12-晚秋胡桃", "8mmPY12-晚秋胡桃"
        ])
        self.assertTrue(all(value == "" for value in output[0][7:9]))
        self.assertEqual(output[0][9], "5+8MM双贴面")  # 打孔工艺列 = 源表工艺列
        self.assertEqual(stats.output_rows, 2)
        self.assertEqual(stats.quantity_sum, 4)
        self.assertEqual(stats.split_source_rows, 1)

    def test_duplicate_layers_are_kept(self) -> None:
        row = [
            "门扇", 2200, 800, 1, 50, "YM-108", "YSM230-3", "ORDER-2",
            "云南省.曲靖市.客户", "4+4mm", "开72锁孔", "逸品",
        ]
        output, _ = transform_rows(HEADERS, [row])
        self.assertEqual([item[17] for item in output], [4, 4])
        self.assertEqual([item[5] for item in output], ["4mm素板", "4mm素板"])
        self.assertEqual([item[21] for item in output], ["4mmYSM230-3", "4mmYSM230-3"])

    def test_missing_process_thickness_leaves_final_thickness_blank(self) -> None:
        row = [
            "门扇", 2490, 830, 1, 50, "YM-083", "YSM230-3", "ORDER-3",
            "广东省.广州市.客户", "复合贴YSM皮装板", "客户特殊要求", "精品包装-S",
        ]
        output, stats = transform_rows(HEADERS, [row])
        self.assertEqual(output[0][5], "YSM230-3")
        self.assertEqual(output[0][17], "")
        self.assertEqual(output[0][21], "YSM230-3")
        self.assertEqual(stats.blank_thickness_orders, ("订单 ORDER-3：YM-083",))

    def test_csv_uses_template_encoding_and_expected_headers(self) -> None:
        row = ["ORDER-1", "单扇", 2200, 800, 2, "8mm颜色", 1]
        row += ["", "", "", "工艺", "", "", "客户", "", "", "品牌", 8]
        row += ["", "", "", "8mm颜色", "", "", "", ""]
        data = write_csv([row])
        self.assertFalse(data.startswith(b"\xef\xbb\xbf"))
        parsed = list(csv.reader(io.StringIO(data.decode("gb18030"))))
        self.assertEqual(parsed[0], OUTPUT_HEADERS)
        self.assertEqual(parsed[1][5], "8mm颜色")
        self.assertEqual(parsed[1][21], "8mm颜色")

    def test_missing_header_is_rejected(self) -> None:
        with self.assertRaisesRegex(ConversionError, "源表缺少字段"):
            transform_rows(HEADERS[:-1], [])

    def test_board_material_follows_process_keywords(self) -> None:
        rows = [
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-1", "重庆市.重庆市.客户", "8mm多层加密", "不开锁孔", "美心蒙迪"],
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-2", "重庆市.重庆市.客户", "5+8MM黑炭晶", "不开锁孔", "美心蒙迪"],
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-3", "重庆市.重庆市.客户", "8mm双贴面", "不开锁孔", "美心蒙迪"],
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-4", "重庆市.重庆市.客户", "8mm黑碳晶", "不开锁孔", "美心蒙迪"],
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-5", "重庆市.重庆市.客户", "PVC系列28MM多层板加密", "不开锁孔", "美心蒙迪"],
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-6", "重庆市.重庆市.客户", "复合贴YSM皮8MM加密多层", "不开锁孔", "美心蒙迪"],
        ]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(output[0][5], "8mm多层加密")
        self.assertEqual([row[5] for row in output[1:3]], ["5mm黑炭晶", "8mm黑炭晶"])
        self.assertEqual(output[3][5], "8mm素板")
        self.assertEqual(output[4][5], "8mm黑碳晶")
        # 多层板加密 / 加密多层 都归一为多层加密，且不触发不一致警告
        self.assertEqual(output[5][5], "28mm多层加密")
        self.assertEqual(output[6][5], "8mm多层加密")
        self.assertEqual(stats.material_mismatch_details, ())
        # 材料描述2 不受影响，仍是 厚度+颜色
        self.assertEqual(output[0][21], "8mmPY12-晚秋胡桃")

    def test_unseen_process_and_mismatch_warnings(self) -> None:
        rows = [
            # 工艺含多层加密但没写厚度 → 材料描述用颜色兜底 → 不一致警告
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-1", "重庆市.重庆市.客户", "多层加密", "不开锁孔", "美心蒙迪"],
            # 没见过的工艺 → 提醒检查
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-2", "重庆市.重庆市.客户", "PET高光板8MM", "不开锁孔", "美心蒙迪"],
            # 已见过且一致 → 无任何警告
            ["门扇", 2240, 810, 1, 50, "N1046单扇", "PY12-晚秋胡桃",
             "ORDER-3", "重庆市.重庆市.客户", "复合贴YSM皮8MM多层加密", "不开锁孔", "美心蒙迪"],
        ]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual(len(stats.material_mismatch_details), 1)
        self.assertIn("ORDER-1", stats.material_mismatch_details[0])
        self.assertEqual(len(stats.new_process_details), 1)
        self.assertIn("ORDER-2", stats.new_process_details[0])
        self.assertIn("PET高光板", stats.new_process_details[0])

    def test_runtime_exclusions_are_merged(self) -> None:
        import tempfile

        runtime_file = Path(tempfile.mkdtemp()) / "runtime_exclusions.json"
        converter.RUNTIME_EXCLUSIONS_FILE = runtime_file
        converter.save_runtime_exclusions({"models": ["YM-999"], "keywords": ["测试关键词"]})

        rows = [
            ["门扇", 2200, 800, 1, 50, "YM999单扇", "YSM230-3", "ORDER-1",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
            ["门扇", 2200, 800, 1, 50, "测试关键词门", "YSM230-3", "ORDER-2",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
            ["门扇", 2200, 800, 1, 50, "YM-061", "YSM230-3", "ORDER-3",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
        ]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual([row[0] for row in output], ["ORDER-3"])
        self.assertEqual(len(stats.excluded_details), 2)
        self.assertIn("YM-999", converter.get_excluded_models())
        self.assertIn("测试关键词", converter.get_excluded_keywords())

    def test_excluded_door_models_are_dropped(self) -> None:
        rows = [
            ["门扇", 2240, 810, 1, 50, "N1054(N1054)单扇", "PY12-晚秋胡桃",
             "ORDER-1", "重庆市.重庆市.客户", "8mm", "不开锁孔", "美心蒙迪"],
            ["门扇", 2200, 800, 2, 50, "YM-062", "YSM230-3", "ORDER-2",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
            ["门扇", 2200, 800, 1, 50, "YM-061", "YSM230-3", "ORDER-3",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
            ["门扇", 2490, 830, 1, 50, "平板无造型", "YSM230-3", "ORDER-4",
             "广东省.广州市.客户", "4mm", "客户特殊要求", "精品包装-S"],
            ["门扇", 2200, 800, 1, 50, "YSM079单扇", "YSM230-3", "ORDER-5",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
            ["门扇", 2200, 800, 1, 50, "ｙｓｍ－０８０ 单扇", "YSM230-3", "ORDER-6",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
            ["门扇", 2200, 800, 1, 50, "铝框门（灰框）", "YSM230-3", "ORDER-7",
             "云南省.曲靖市.客户", "4mm", "开72锁孔", "逸品"],
        ]
        output, stats = transform_rows(HEADERS, rows)

        self.assertEqual([row[0] for row in output], ["ORDER-3", "ORDER-4"])
        self.assertEqual(stats.source_rows, 2)
        self.assertEqual(stats.excluded_details, (
            "订单 ORDER-1：N1054(N1054)单扇",
            "订单 ORDER-2：YM-062",
            "订单 ORDER-5：YSM079单扇",
            "订单 ORDER-6：ｙｓｍ－０８０ 单扇",
            "订单 ORDER-7：铝框门（灰框）（关键词剔除）",
        ))


if __name__ == "__main__":
    unittest.main()
