# -*- coding: utf-8 -*-
"""群里那条消息必须短：只有「单号 + 改了什么」（+ 要核对哪几格）。

以前把逐格长说明整段塞进 warnings，群里一排十几行，车间不看。
这里盯的就是摘要本身：类别顺序、机械动作不带数字、超长截断、必看项单列。
"""
import sys
import unittest
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC_DIR))

from make_change_notice import build_brief, order_hint  # noqa: E402


def cell(category, sheet="1", coord="L23", old="a", new="b"):
    return {"category": category, "sheet": sheet, "coord": coord, "old": old, "new": new}


HINT = order_hint("B2611-8079陈瑞利补单")


class ChangeBriefTests(unittest.TestCase):
    def test_order_number_and_changes_on_one_line(self) -> None:
        items = [cell("单价折扣")] * 213 + [cell("木箱包装行清空")] + [cell("备注清理（删掉结算/折扣金额）")] * 3
        line, review = build_brief(items, {}, HINT)
        self.assertEqual(line.splitlines()[0], "B2611-8079 · 删木箱行1处 · 删备注3处 · 单价8折")
        self.assertEqual(review, "")

    def test_note_deletion_spells_out_the_original_text(self) -> None:
        """删备注必须把原文写全：只报「删备注 1 处」等于什么都没说。"""
        items = [cell("备注清理（删掉结算/折扣金额）", coord="D32",
                      old="王董同意打9.2折，折后：13230元，折后：1150元", new=None)]
        line, _ = build_brief(items, {}, HINT)
        self.assertEqual(line.splitlines(), [
            "B2611-8079 · 删备注1处",
            "🗒 删备注（1页D32）：「王董同意打9.2折，折后：13230元，折后：1150元」",
        ])

    def test_note_partial_deletion_keeps_only_removed_sentences(self) -> None:
        items = [cell("备注清理（删掉结算/折扣金额）", coord="E8",
                      old="此单打8折。折后：9800元。送货到工地。", new="送货到工地。")]
        line, _ = build_brief(items, {}, HINT)
        self.assertIn("「此单打8折；折后：9800元」", line)
        self.assertNotIn("送货到工地", line)

    def test_each_distinct_note_gets_its_own_line(self) -> None:
        items = [cell("备注清理（删掉结算/折扣金额）", coord="D32", old="打9.2折，折后：13230元", new=None),
                 cell("备注清理（删掉结算/折扣金额）", sheet="2", coord="E8",
                      old="优惠400元，需支付8800元", new=None)]
        lines = build_brief(items, {}, HINT)[0].splitlines()
        self.assertEqual(lines[1], "🗒 删备注（1页D32）：「打9.2折，折后：13230元」")
        self.assertEqual(lines[2], "🗒 删备注（2页E8）：「优惠400元，需支付8800元」")

    def test_same_note_on_two_sheets_merges_places(self) -> None:
        items = [cell("备注清理（删掉结算/折扣金额）", sheet=s, old="打9.2折，折后：13230元", new=None)
                 for s in ("1", "2")]
        lines = build_brief(items, {}, HINT)[0].splitlines()
        self.assertEqual(lines[1], "🗒 删备注（1页L23、2页L23）：「打9.2折，折后：13230元」")

    def test_too_many_notes_get_a_pointer_not_silence(self) -> None:
        items = [cell("备注清理（删掉结算/折扣金额）", coord=f"Q{i}", old=f"备注{i}，折后：{i}00元", new=None)
                 for i in range(9)]
        lines = build_brief(items, {}, HINT)[0].splitlines()
        self.assertEqual(len(lines), 1 + 8 + 1)
        self.assertIn("另有 1 处备注，原文见改动清单", lines[-1])

    def test_mechanical_bulk_actions_carry_no_count(self) -> None:
        """单价 8 折、公式定值、页号校正说"多少处"对车间没意义。"""
        items = ([cell("单价折扣")] * 200
                 + [cell("公式固化为数值（结果不变）")] * 158
                 + [cell("订单编号后缀更正")] * 4
                 + [cell("木箱包装行清空")])
        line, _ = build_brief(items, {}, HINT)
        self.assertEqual(line, "B2611-8079 · 删木箱行1处 · 公式定值 · 页号校正 · 单价8折")

    def test_internal_follow_cells_are_hidden(self) -> None:
        """连带更新是内部噪音，一个字都别出现在群里。"""
        items = [cell("连带更新（跟随被引用单元格）")] * 30
        line, review = build_brief(items, {}, HINT)
        self.assertEqual(line, "B2611-8079 · 没改动")
        self.assertEqual(review, "")

    def test_review_line_lists_cells(self) -> None:
        items = [cell("单价折扣"), cell("内容改写", sheet="2", coord="T9")]
        line, review = build_brief(items, {}, HINT)
        self.assertEqual(line, "B2611-8079 · 内容变1处 · 单价8折")
        self.assertEqual(review, "⚠ 核对1处：2页T9")

    def test_review_line_caps_at_four_spots(self) -> None:
        items = [cell("内容改写", coord=f"Q{i}") for i in range(1, 8)]
        _, review = build_brief(items, {}, HINT)
        self.assertTrue(review.startswith("⚠ 核对7处：1页Q1、1页Q2、1页Q3、1页Q4"))
        self.assertIn("等7处", review)

    def test_too_many_categories_collapse(self) -> None:
        items = [cell("木箱包装行清空"), cell("备注清理（删掉结算/折扣金额）"),
                 cell("工分补填"), cell("内容改写"), cell("新增内容"), cell("删除内容")]
        line, _ = build_brief(items, {}, HINT)
        self.assertEqual(line.splitlines()[0],
                         "B2611-8079 · 内容变1处 · 新增1处 · 删除1处 · 删木箱行1处 · 等2类")

    def test_no_changes(self) -> None:
        self.assertEqual(build_brief([], {}, HINT), ("B2611-8079 · 没改动", ""))

    def test_missing_order_number_falls_back_to_body(self) -> None:
        line, _ = build_brief([cell("木箱包装行清空")], {}, "")
        self.assertEqual(line, "删木箱行1处")


if __name__ == "__main__":
    unittest.main()
