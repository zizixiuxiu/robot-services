import sys
import tempfile
import unittest
from base64 import b64decode
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image


SRC_DIR = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(SRC_DIR))

from make_workshop_order import (  # noqa: E402
    TransformStats,
    _inject_cached_values,
    clean_payment_amount_text,
    collect_formula_cells,
    fill_workpoint_formulas,
    find_existing_order_number,
    fit_long_note_rows,
    freeze_external_links,
    infer_order_number,
    is_hardware_sheet_name,
    is_hunyou_color,
    remove_payment_amount_info,
    restore_color_block_merge,
    transform,
)
from openpyxl.workbook.external_link.external import (  # noqa: E402
    ExternalBook,
    ExternalCell,
    ExternalLink,
    ExternalRow,
    ExternalSheetData,
    ExternalSheetDataSet,
    ExternalSheetNames,
)


PNG_1X1 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB"
    "/6X7x2sAAAAASUVORK5CYII="
)


class WorkshopOrderRuleTests(unittest.TestCase):
    def test_hunyou_color_allows_leading_brackets(self) -> None:
        self.assertTrue(is_hunyou_color("（混油）Y05-S月影灰"))
        self.assertTrue(is_hunyou_color("(混油) Y05-S月影灰"))
        self.assertFalse(is_hunyou_color("天然木皮：白橡"))

    def test_note_cleanup_keeps_after_sales_sentence_and_strips_price(self) -> None:
        source = "玻璃运输易损坏，不进入公司售后！附色卡有色差，具体以出厂实物为准，不进入公司售后！92折优惠：972元"
        self.assertEqual(
            remove_payment_amount_info(source),
            "玻璃运输易损坏，不进入公司售后！附色卡有色差，具体以出厂实物为准，不进入公司售后！",
        )

    def test_payment_amount_cleanup_strips_discount_segment_after_after_sales(self) -> None:
        source = (
            "此单黄杨原木修红樱桃颜色纹理不同修色有色差不进入公司售后！工艺不同，以实物为准,不进入公司售后！"
            "               此单王董同意板材95折优惠，折后31197元，优惠1642元！"
        )
        self.assertEqual(
            remove_payment_amount_info(source),
            "此单黄杨原木修红樱桃颜色纹理不同修色有色差不进入公司售后！工艺不同，以实物为准,不进入公司售后！",
        )

    def test_payment_amount_cleanup_keeps_after_sales_before_discount_approval(self) -> None:
        prefix = "此单工艺不同，以实物为准，不进入公司售后！"
        examples = (
            prefix + "此单李总同意板材95折优惠，折后31197元，优惠1642元！",
            prefix + "此单张总确认8折，优惠100元！",
            prefix + "此单赵董同意折后31197元！",
        )
        for source in examples:
            self.assertEqual(remove_payment_amount_info(source), prefix)

    def test_payment_amount_cleanup_clears_price_only_fragments(self) -> None:
        self.assertIsNone(clean_payment_amount_text("银镜320元/平").value)
        self.assertIsNone(clean_payment_amount_text("样品打8折、折后：119810元，优惠：29953元").value)

    def test_payment_amount_cleanup_keeps_page_count(self) -> None:
        self.assertEqual(remove_payment_amount_info("此单共2页，92折优惠：972元"), "此单共2页")

    def test_infer_order_number_from_filename(self) -> None:
        path = Path("S2606-4106阮国政（西安越秀铁建樽越）.xlsx")
        self.assertEqual(infer_order_number(path), "S2606-4106")

    def test_existing_header_order_number_wins_over_filename(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "1"
        ws["P2"] = "订单编号"
        ws["Q2"] = "S2606-6111-2"
        ws["R2"] = -1

        self.assertEqual(find_existing_order_number(wb), "S2606-6111-2")

    def test_hardware_sheet_name(self) -> None:
        self.assertTrue(is_hardware_sheet_name("五金-1"))
        self.assertTrue(is_hardware_sheet_name("订单五金汇总"))
        self.assertFalse(is_hardware_sheet_name("1"))

    def test_transform_preserves_embedded_images(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            image_path = tmpdir / "image.png"
            image_path.write_bytes(b64decode(PNG_1X1))

            source_path = tmpdir / "S2606-0001-image.xlsx"
            output_path = tmpdir / "S2606-0001-image下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["A1"] = "测试"
            ws.add_image(Image(str(image_path)), "S8")
            wb.save(source_path)

            transform(source_path, output_path, 0.85, "auto")

            with ZipFile(output_path) as zf:
                media = [name for name in zf.namelist() if name.startswith("xl/media/") and not name.endswith("/")]
            self.assertEqual(len(media), 1)

    def test_transform_preserves_existing_header_order_number(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2606-6111-sample.xlsx"
            output_path = tmpdir / "S2606-6111-sample下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["P2"] = "订单编号"
            ws["Q2"] = "S2606-6111-2"
            ws["R2"] = -1
            wb.save(source_path)

            transform(source_path, output_path, 0.85, "auto")

            out = load_workbook(output_path)
            self.assertEqual(out["1"]["Q2"].value, "S2606-6111-2")
            self.assertEqual(out["1"]["R2"].value, -1)

    def test_transform_warns_when_no_cached_value_written(self) -> None:
        """算不出结果时必须在群里看见，否则只有飞书网页版会露馅。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2606-7000-sample.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["P2"] = "订单编号"
            ws["Q2"] = "S2606-7000"
            ws["E8"] = 100
            ws["F8"] = 200
            ws["H8"] = 1
            ws["J8"] = "=E8/1000*F8/1000*H8"
            wb.save(source_path)

            with patch("make_workshop_order.embed_cached_values", return_value=(0, 0)):
                stats = transform(source_path, tmpdir / "out.xlsx", 0.8, "auto")

            warnings = " ".join(stats["warnings"])
            self.assertIn("公式结果全部没算出来", warnings)
            self.assertIn("飞书网页版", warnings)

    def test_transform_warns_when_only_part_written_back(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2606-7001-sample.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["P2"] = "订单编号"
            ws["Q2"] = "S2606-7001"
            ws["J8"] = "=1+1"
            ws["J9"] = "=2+2"
            wb.save(source_path)

            with patch("make_workshop_order.embed_cached_values", return_value=(2, 1)):
                stats = transform(source_path, tmpdir / "out.xlsx", 0.8, "auto")

            self.assertIn("只写回 1/2", " ".join(stats["warnings"]))

    def test_transform_quiet_when_all_written_back(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2606-7002-sample.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["P2"] = "订单编号"
            ws["Q2"] = "S2606-7002"
            ws["J8"] = "=1+1"
            wb.save(source_path)

            with patch("make_workshop_order.embed_cached_values", return_value=(1, 1)):
                stats = transform(source_path, tmpdir / "out.xlsx", 0.8, "auto")

            self.assertEqual(stats["warnings"], [])

    def test_existing_header_order_number_with_urgent_prefix(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "1"
        ws["P2"] = "订单编号"
        ws["Q2"] = "加急S2610-6059-3"

        self.assertEqual(find_existing_order_number(wb), "加急S2610-6059-3")

    def test_transform_keeps_urgent_prefix_and_sub_number(self) -> None:
        """「加急S2610-6059-3」不能被文件名推出来的 S2610-6059 覆盖。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2610-6059 直营店(陈樱）.xlsx"
            output_path = tmpdir / "S2610-6059 直营店(陈樱）下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["P2"] = "订单编号"
            ws["Q2"] = "加急S2610-6059-3"
            ws["R2"] = -1
            wb.save(source_path)

            transform(source_path, output_path, 0.85, "auto")

            out = load_workbook(output_path)
            self.assertEqual(out["1"]["Q2"].value, "加急S2610-6059-3")
            self.assertEqual(out["1"]["R2"].value, -1)

    def test_transform_does_not_create_missing_order_suffix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2607-7506-sample.xlsx"
            output_path = tmpdir / "S2607-7506-sample下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            ws["P2"] = "订单编号"
            ws["Q2"] = "S2607-7506"
            ws["R2"] = None
            wb.save(source_path)

            transform(source_path, output_path, 0.85, "auto")

            out = load_workbook(output_path)
            self.assertEqual(out["1"]["Q2"].value, "S2607-7506")
            self.assertIsNone(out["1"]["R2"].value)

    def test_fit_long_note_rows_replaces_space_padding_with_line_break(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.merge_cells("D25:R25")
        ws["D25"] = "此单黄杨原木修红樱桃颜色纹理不同修色有色差不进入公司售后！        此单王董同意板材"
        ws.row_dimensions[25].height = 20

        fit_long_note_rows(ws)

        self.assertIn("\n", ws["D25"].value)
        self.assertTrue(ws["D25"].alignment.wrap_text)
        self.assertGreater(ws.row_dimensions[25].height, 20)

    def test_fit_long_note_rows_skips_tall_multi_row_merged_note(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.merge_cells("A15:C18")
        ws["A15"] = "图文说明:请车间注意在外包装上写上，内有玻璃，勿压!!!!!"
        for row_idx in range(15, 19):
            ws.row_dimensions[row_idx].height = 25

        fit_long_note_rows(ws)

        # 合并区域 4 行合计 100pt 已足够，首行不应被加高
        self.assertEqual(ws.row_dimensions[15].height, 25)

    def test_fit_long_note_rows_tops_up_short_multi_row_merged_note(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.merge_cells("A15:C16")
        ws["A15"] = "图文说明:" + "请车间注意在外包装上写上，内有玻璃，勿压!!!!!" * 5
        for row_idx in range(15, 17):
            ws.row_dimensions[row_idx].height = 25

        fit_long_note_rows(ws)

        # 合并区域合计高度不够时，差值补到首行，而不是整段高度
        self.assertGreater(ws.row_dimensions[15].height, 25)
        self.assertLessEqual(ws.row_dimensions[15].height + 25, 180)

    def _quote_sheet_headers(self, ws) -> None:
        ws["C7"] = "产品名称"
        ws["K7"] = "单价（元）"
        ws["L7"] = "小计（元）"
        ws["M7"] = "工分"
        ws["N7"] = "颜色"
        ws.merge_cells("N7:O7")

    def test_blank_row_delete_keeps_spanning_color_merge(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2607-4003-color.xlsx"
            output_path = tmpdir / "S2607-4003-color下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            self._quote_sheet_headers(ws)
            ws["C8"] = "左立板"
            ws["K8"] = 320
            ws["L8"] = 100
            ws["N8"] = "混油附色卡（同免漆月影灰X）（请车间多配一瓶油漆，随货一起）"
            ws.merge_cells("N8:O9")
            ws["C10"] = "合计"
            ws["L10"] = "=SUM(L8:L8)"
            wb.save(source_path)

            transform(source_path, output_path, 0.8, "auto")

            out = load_workbook(output_path)["1"]
            merged = {str(rng) for rng in out.merged_cells.ranges}
            self.assertIn("N8:O8", merged)
            self.assertEqual(
                out["N8"].value,
                "混油附色卡（同免漆月影灰X）（请车间多配一瓶油漆，随货一起）",
            )

    def test_restore_color_block_merge_fills_quote_style_block(self) -> None:
        wb = Workbook()
        ws = wb.active
        self._quote_sheet_headers(ws)
        ws["C8"] = "左立板"
        ws["C9"] = "平板"
        ws["C10"] = "右立板"
        ws["N8"] = "混油附色卡（同免漆月影灰X）"
        ws.merge_cells("N8:O8")
        ws.merge_cells("N9:O9")
        ws.merge_cells("N10:O10")
        ws["C11"] = "合计"

        restored = restore_color_block_merge(ws, 7, 3, 14)

        self.assertTrue(restored)
        merged = {str(rng) for rng in ws.merged_cells.ranges}
        self.assertIn("N8:O10", merged)
        self.assertNotIn("N8:O8", merged)
        self.assertEqual(ws["N8"].value, "混油附色卡（同免漆月影灰X）")
        self.assertIsNone(ws["N9"].value)

    def test_restore_color_block_merge_covers_merged_product_name_rows(self) -> None:
        wb = Workbook()
        ws = wb.active
        self._quote_sheet_headers(ws)
        ws["C8"] = "造型墙板"
        ws["C9"] = "花格"
        ws["C10"] = "条子"
        ws.merge_cells("C10:C13")
        ws["E10"] = 2710
        ws["E11"] = 2710
        ws["E12"] = 2540
        ws["E13"] = 2400
        ws["N8"] = "红胡桃木皮修ZKP-3103封闭漆附色卡"
        ws.merge_cells("N8:O8")
        ws["C14"] = "木箱包装"
        ws["C15"] = "合计"

        restored = restore_color_block_merge(ws, 7, 3, 14)

        self.assertTrue(restored)
        merged = {str(rng) for rng in ws.merged_cells.ranges}
        self.assertIn("N8:O13", merged)
        self.assertEqual(ws["N8"].value, "红胡桃木皮修ZKP-3103封闭漆附色卡")

    def test_restore_color_block_merge_covers_size_continuation_rows(self) -> None:
        wb = Workbook()
        ws = wb.active
        self._quote_sheet_headers(ws)
        ws["C8"] = "造型墙板"
        ws["C9"] = "条子"
        ws["E9"] = 2710
        ws["E10"] = 2710
        ws["E11"] = 2540
        ws["E12"] = 2400
        ws["N8"] = "红胡桃木皮修ZKP-3103封闭漆附色卡"
        ws.merge_cells("N8:O8")
        ws["C13"] = "合计"

        restored = restore_color_block_merge(ws, 7, 3, 14)

        self.assertTrue(restored)
        merged = {str(rng) for rng in ws.merged_cells.ranges}
        self.assertIn("N8:O12", merged)

    def test_restore_color_block_merge_follows_material_merge(self) -> None:
        wb = Workbook()
        ws = wb.active
        self._quote_sheet_headers(ws)
        ws["P7"] = "材质"
        ws["C8"] = "造型墙板"
        ws["C9"] = "花格"
        ws["N8"] = "红胡桃木皮修ZKP-3103封闭漆附色卡"
        ws.merge_cells("N8:O8")
        ws["P8"] = "密度板"
        ws.merge_cells("P8:P13")
        ws["C14"] = "合计"

        restored = restore_color_block_merge(ws, 7, 3, 14)

        self.assertTrue(restored)
        merged = {str(rng) for rng in ws.merged_cells.ranges}
        self.assertIn("N8:O13", merged)

    def test_restore_color_block_merge_keeps_segmented_accent_colors(self) -> None:
        """分段颜色（叉色）已存在时，禁止合并成单一颜色块。"""
        wb = Workbook()
        ws = wb.active
        self._quote_sheet_headers(ws)
        ws["C8"] = "墙板A"
        ws["C11"] = "柜体B"
        ws["C12"] = "墙板C"
        ws["N8"] = "混油肤感附色卡（雪山白）"
        ws.merge_cells("N8:O10")
        ws["N11"] = "混油肤感附色卡（雪山白）/见图叉色混油肤感附色卡（幻夜黑）"
        ws.merge_cells("N11:O11")
        ws["N12"] = "混油肤感附色卡（雪山白）"
        ws.merge_cells("N12:O15")
        ws["C16"] = "合计"

        restored = restore_color_block_merge(ws, 7, 3, 14)

        self.assertFalse(restored)
        merged = {str(rng) for rng in ws.merged_cells.ranges}
        self.assertIn("N8:O10", merged)
        self.assertIn("N11:O11", merged)
        self.assertIn("N12:O15", merged)
        self.assertNotIn("N8:O15", merged)
        self.assertEqual(ws["N11"].value, "混油肤感附色卡（雪山白）/见图叉色混油肤感附色卡（幻夜黑）")

    def test_transform_keeps_color_merge_across_strip_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2608-9007-color.xlsx"
            output_path = tmpdir / "S2608-9007-color下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            self._quote_sheet_headers(ws)
            ws["P7"] = "材质"
            ws["C8"] = "造型墙板"
            ws["K8"] = 430
            ws["L8"] = 100
            ws["C9"] = "花格"
            ws["K9"] = 780
            ws["L9"] = 200
            ws["C10"] = "条子"
            ws.merge_cells("C10:C13")
            for row, height in ((10, 2710), (11, 2710), (12, 2540), (13, 2400)):
                ws.cell(row, 5).value = height
                ws.cell(row, 8).value = 2
                ws.cell(row, 11).value = 23
                ws.cell(row, 12).value = 40
            ws["N8"] = "红胡桃木皮修ZKP-3103封闭漆附色卡"
            ws.merge_cells("N8:O13")
            ws["P8"] = "密度板"
            ws.merge_cells("P8:P13")
            ws["C14"] = "木箱包装"
            ws["L14"] = 150
            ws["C15"] = "合计"
            wb.save(source_path)

            transform(source_path, output_path, 0.8, "auto")

            out = load_workbook(output_path)["1"]
            merged = {str(rng) for rng in out.merged_cells.ranges}
            self.assertIn("N8:O13", merged)
            self.assertEqual(out["N8"].value, "红胡桃木皮修ZKP-3103封闭漆附色卡")
            self.assertIn("P8:P13", merged)

    def test_merged_subtotal_block_workpoints_count_once(self) -> None:
        """门套整组小计纵向合并（按套计价）时，工分只在首行算一次，其余行为 0。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2608-6049-doorset.xlsx"
            output_path = tmpdir / "S2608-6049-doorset下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "6"
            self._quote_sheet_headers(ws)
            ws["C8"] = "门洞尺寸"
            ws["K8"] = 4478
            ws["L8"] = 4478
            ws.merge_cells("K8:K16")
            ws.merge_cells("L8:L16")
            for row, name in (
                (9, "侧套板"),
                (10, "顶套板"),
                (11, "50套线"),
                (12, "门扇"),
                (13, "门扇"),
                (14, "门扇"),
                (15, "门扇"),
                (16, "门扇"),
            ):
                ws.cell(row, 3).value = name
            # 模拟错误：合并块内多行都写了 =L8/100
            ws["M8"] = "=L8/100"
            ws["M9"] = "=L8/100"
            # 非合并产品行，应正常写工分
            ws["C17"] = "踢脚线"
            ws["L17"] = 200
            ws["C18"] = "合计"
            ws["M18"] = "=SUM(M8:M17)"
            wb.save(source_path)

            transform(source_path, output_path, 1.0, "auto")

            out = load_workbook(output_path)["6"]
            self.assertEqual(out["M8"].value, "=L8/100")
            for row in range(9, 17):
                self.assertEqual(out.cell(row, 13).value, 0, f"M{row} should be 0")
            self.assertEqual(out["M17"].value, "=L17/100")

    def test_fill_workpoint_zeros_merged_subtotal_directly(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws["C7"] = "产品名称"
        ws["L7"] = "小计（元）"
        ws["M7"] = "工分"
        ws["C8"] = "门洞尺寸"
        ws["L8"] = 4478
        ws.merge_cells("L8:L16")
        ws["M8"] = "=L8/100"
        ws["M9"] = "=L8/100"
        ws["C17"] = "合计"

        filled = fill_workpoint_formulas(ws, 7, 3, 12, 13)

        self.assertGreater(filled, 0)
        self.assertEqual(ws["M8"].value, "=L8/100")
        self.assertIn("0.00", str(ws["M8"].number_format))
        for row in range(9, 17):
            self.assertEqual(ws.cell(row, 13).value, 0)

    def test_fill_workpoint_workpoint_column_merged_vertically(self) -> None:
        """工分列本身纵向合并时，工分应写在合并块左上角 M8。"""
        wb = Workbook()
        ws = wb.active
        ws["C7"] = "产品名称"
        ws["L7"] = "小计（元）"
        ws["M7"] = "工分"
        ws["C8"] = "门洞尺寸"
        ws["L8"] = 4478
        ws.merge_cells("L8:L16")
        ws.merge_cells("M8:M16")
        ws["C17"] = "合计"

        filled = fill_workpoint_formulas(ws, 7, 3, 12, 13)

        self.assertGreater(filled, 0)
        self.assertEqual(ws["M8"].value, "=L8/100")
        self.assertIn("0.00", str(ws["M8"].number_format))
        for row in range(9, 17):
            self.assertIsInstance(ws.cell(row, 13), MergedCell, f"M{row} should be merged")

    def test_fill_workpoint_multi_column_merge_does_not_write_to_wrong_column(self) -> None:
        """工分列被跨列合并且左上角不在工分列时，不应把工分公式写到别的列。"""
        wb = Workbook()
        ws = wb.active
        ws["C7"] = "产品名称"
        ws["L7"] = "小计（元）"
        ws["M7"] = "工分"
        ws["C8"] = "门洞尺寸"
        ws["L8"] = 4478
        ws.merge_cells("L8:L16")
        # L:M 跨列合并，M8-M16 都是 MergedCell，左上角在 L 列
        ws.merge_cells("L8:M16")
        ws["C17"] = "合计"

        filled = fill_workpoint_formulas(ws, 7, 3, 12, 13)

        # 工分列无法独立写入工分公式；但合计行仍可能写 SUM
        for row in range(8, 17):
            self.assertIsInstance(ws.cell(row, 13), MergedCell)
        # 确认 L 列小计值未被破坏
        self.assertEqual(ws["L8"].value, 4478)

    def test_workpoint_keeps_decimal_not_integer_format(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws["C7"] = "产品名称"
        ws["L7"] = "小计（元）"
        ws["M7"] = "工分"
        ws["C8"] = "门洞"
        ws["L8"] = 3264
        ws.merge_cells("L8:L12")
        ws["M8"] = "=L8/100"
        ws["M8"].number_format = "0;[Red]0"
        ws["C13"] = "合计"

        fill_workpoint_formulas(ws, 7, 3, 12, 13)

        self.assertEqual(ws["M8"].value, "=L8/100")
        self.assertIn("0.00", str(ws["M8"].number_format))

    def test_long_note_mentioning_wood_box_does_not_wipe_page_total(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            source_path = tmpdir / "S2607-4068-sample.xlsx"
            output_path = tmpdir / "S2607-4068-sample下车间.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "1"
            # 表头：产品/单价/小计/工分
            ws["C7"] = "产品名称"
            ws["K7"] = "单价（元）"
            ws["L7"] = "小计（元）"
            ws["M7"] = "工分"
            # 一条普通产品行
            ws["C8"] = "80欧式线条"
            ws["K8"] = 32
            ws["L8"] = "=I8*K8"
            # 木箱包装行（短标签，应被清除金额）
            ws["C9"] = "木箱包装"
            ws["D9"] = "此单共打4个木箱包装"
            ws["L9"] = "=150*4"
            # 图文说明长备注里提到“木箱包装”，同行的页合计公式必须保留
            ws["A10"] = "图文说明:玻门单独打木箱包装，与木门单JF26-07-28-70012一起油漆出货"
            ws["L10"] = "=+SUM(L8:L9)"
            wb.save(source_path)

            transform(source_path, output_path, 0.85, "auto")

            out = load_workbook(output_path)
            self.assertIsNone(out["1"]["L9"].value)
            self.assertEqual(out["1"]["L10"].value, "=+SUM(L8:L9)")

    def test_freeze_external_links_replaces_formula_and_drops_link(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "1"
        ws["E3"] = "='[1]1'!Q5"
        ws["A27"] = "='[1]汇总'!D23"
        book = ExternalBook(
            sheetNames=ExternalSheetNames(sheetName=["1", "汇总"]),
            sheetDataSet=ExternalSheetDataSet(sheetData=[
                ExternalSheetData(sheetId=0, row=[ExternalRow(r=5, cell=[ExternalCell(r="Q5", v="-1")])]),
                ExternalSheetData(sheetId=1, row=[ExternalRow(r=23, cell=[ExternalCell(r="D23", t="str", v="此单共页")])]),
            ]),
        )
        link = ExternalLink(externalBook=book)
        wb._external_links = [link]
        stats = TransformStats(order_type="tiepi")

        frozen = freeze_external_links(wb, stats)

        self.assertEqual(frozen, 2)
        self.assertEqual(ws["E3"].value, -1)
        self.assertEqual(ws["A27"].value, "此单共页")
        self.assertEqual(wb._external_links, [])


class CachedValueTests(unittest.TestCase):
    """openpyxl 只写公式不写结果，飞书网页版预览不会重算 → 公式列整列空白。"""

    def _make_formula_book(self, path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "1"
        ws["E8"] = 970
        ws["F8"] = 346
        ws["H8"] = 1
        ws["J8"] = "=E8/1000*F8/1000*H8"
        ws["K8"] = "=340+60"
        ws["L8"] = "=J8*K8"
        ws["M8"] = "=L8/100"
        ws["N8"] = "混油纯黑水性漆"
        ws["H12"] = "=SUM(H8:H8)"
        summary = wb.create_sheet("汇总")
        summary["A3"] = "1"
        summary["B3"] = "='1'!L8"
        wb.save(path)

    def test_inject_writes_cached_values_without_touching_other_cells(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "order.xlsx"
            self._make_formula_book(path)

            before = load_workbook(path, data_only=True)
            self.assertIsNone(before["1"]["L8"].value)  # 缓存值缺失：预览器看到的就是空白
            refs = collect_formula_cells(load_workbook(path))

            filled = _inject_cached_values(
                path,
                {
                    "1": {
                        "J8": ("0.33562", ""),
                        "K8": ("400", ""),
                        "L8": ("134.248", ""),
                        "M8": ("1.34248", ""),
                        "H12": ("1", ""),
                    },
                    "汇总": {"B3": ("134.248", "")},
                },
            )

            self.assertEqual(filled, 6)
            after = load_workbook(path, data_only=True)
            self.assertAlmostEqual(after["1"]["L8"].value, 134.248)
            self.assertAlmostEqual(after["1"]["M8"].value, 1.34248)
            self.assertAlmostEqual(after["汇总"]["B3"].value, 134.248)
            # 普通单元格和公式本身都不能动
            self.assertEqual(after["1"]["N8"].value, "混油纯黑水性漆")
            self.assertEqual(load_workbook(path)["1"]["L8"].value, "=J8*K8")
            self.assertEqual(set(refs), {"1", "汇总"})

    def test_inject_keeps_string_result_typed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "order.xlsx"
            self._make_formula_book(path)

            self.assertEqual(
                _inject_cached_values(path, {"1": {"H12": ("此单共2页", "str")}}), 1
            )

            ws = load_workbook(path, data_only=True)["1"]
            self.assertEqual(ws["H12"].value, "此单共2页")
            raw = ZipFile(path).read("xl/worksheets/sheet1.xml").decode()
            self.assertIn('<c r="H12" t="str">', raw)
            self.assertIn("<v>此单共2页</v>", raw)


if __name__ == "__main__":
    unittest.main()
