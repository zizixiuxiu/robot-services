#!/usr/bin/env python3
"""
工厂版报价单生成器（开发版 v2）
- 隐藏所有价格相关的列和行，保留所有单元格格式（包括边框）
- 图文说明行：不隐藏，去掉金额列，只保留"图文说明:"文本
- 下单日期：通过 order_date 参数传入替换原日期
"""

import xlrd
from xlutils.copy import copy
import xlwt
from xlwt import Utils
import os
import re
import copy as pycopy


def _patch_xlwt_numeric_sheet_names():
    """Make xlwt resolve quoted numeric sheet names like '1' as sheet names."""
    if getattr(xlwt.Workbook, "_numeric_sheet_name_patch", False):
        return

    original_convert_sheetindex = xlwt.Workbook.convert_sheetindex

    def convert_sheetindex(self, strg_ref, n_sheets):
        worksheet_idx_from_name = getattr(self, "_Workbook__worksheet_idx_from_name", {})
        if isinstance(strg_ref, str):
            sheet_idx = worksheet_idx_from_name.get(strg_ref.lower())
            if sheet_idx is not None:
                return sheet_idx
        return original_convert_sheetindex(self, strg_ref, n_sheets)

    xlwt.Workbook.convert_sheetindex = convert_sheetindex
    xlwt.Workbook._numeric_sheet_name_patch = True


_patch_xlwt_numeric_sheet_names()

# 标准文件列宽/行高（来自 B2605-4117陈时伟（融汇11-1-202）_工厂版(1).xls Sheet 1）
NORMAL_STANDARD_COL_WIDTHS = {
    0: 1536, 1: 2272, 2: 3679, 3: 3647, 4: 2976, 5: 2496, 6: 928,
    7: 1280, 8: 1055, 9: 1631, 10: 1216, 11: 3935, 12: 1984,
    13: 1984, 14: 2399, 15: 1024, 16: 1119, 17: 1119,
}

NORMAL_STANDARD_ROW_HEIGHTS = {
    0: 600, 1: 640, 2: 560, 3: 760, 4: 540, 5: 540, 6: 560,
    7: 440, 8: 440, 9: 440, 10: 440, 11: 720,
    12: 8190, 13: 8190, 14: 8190, 15: 8190,
    16: 480, 17: 600, 18: 600, 19: 480, 20: 540,
    21: 360, 22: 480, 23: 480, 24: 440, 25: 480, 26: 439,
}

# 标准文件列宽/行高（来自 B2605-4117陈时伟（融汇11-1-202）_工厂版(1).xls Sheet 1五）
WU_STANDARD_COL_WIDTHS = {
    0: 1133, 1: 438, 2: 2267, 3: 438, 4: 914, 5: 987, 6: 1865,
    7: 1133, 8: 2011, 9: 2816, 10: 1572, 11: 1133, 12: 219,
    13: 1170, 14: 1572, 15: 2560, 16: 1865, 17: 2450,
    18: 4169, 19: 1280, 20: 4754,
}

WU_STANDARD_ROW_HEIGHTS = {
    0: 420, 1: 237, 2: 458, 3: 469, 4: 559,
    5: 559, 6: 559, 7: 559, 8: 559, 9: 559,
    10: 559, 11: 559, 12: 559, 13: 559, 14: 559,
    15: 559, 16: 559, 17: 559, 18: 559, 19: 559,
    20: 559, 21: 559, 22: 559, 23: 559,
    24: 8190, 25: 810, 26: 469, 27: 458,
}

WU_LAYOUT_TEMPLATE_PATHS = [
    os.path.join(os.path.dirname(__file__), "wu_layout_template.xls"),
    r"C:\Users\Administrator\Documents\Codex\2026-05-26\1-xlwt-formula-sheet-1-sheet\wu_wrap_layout_test.xls",
    "/mnt/c/Users/Administrator/Documents/Codex/2026-05-26/1-xlwt-formula-sheet-1-sheet/wu_wrap_layout_test.xls",
]
WU_LAYOUT_TEMPLATE_SHEET = '1五'

NORMAL_LAYOUT_TEMPLATE_PATHS = [
    os.path.join(os.path.dirname(__file__), "normal_layout_template.xls"),
]
NORMAL_LAYOUT_TEMPLATE_SHEET = '1'

# 线型映射：xlrd -> xlwt
LINE_STYLE_MAP = {
    0: xlwt.Borders.NO_LINE,
    1: xlwt.Borders.THIN,
    2: xlwt.Borders.MEDIUM,
    3: xlwt.Borders.DASHED,
    4: xlwt.Borders.DOTTED,
    5: xlwt.Borders.THICK,
    6: xlwt.Borders.DOUBLE,
    7: xlwt.Borders.HAIR,
}

# 样式缓存 - 按 (xf_index, row, col) 缓存，确保每个单元格独立样式
style_cache = {}


def get_cell_style(book, sheet, rowx, colx):
    """获取单元格的完整样式（每个单元格独立，避免空单元格共享样式导致边框丢失）"""
    xf_index = sheet.cell_xf_index(rowx, colx)
    cache_key = (xf_index, rowx, colx)
    
    if cache_key in style_cache:
        return style_cache[cache_key]
    
    xf = book.xf_list[xf_index]
    style = xlwt.XFStyle()
    
    # 复制边框
    borders = xlwt.Borders()
    borders.left = LINE_STYLE_MAP.get(xf.border.left_line_style, xlwt.Borders.NO_LINE)
    borders.right = LINE_STYLE_MAP.get(xf.border.right_line_style, xlwt.Borders.NO_LINE)
    borders.top = LINE_STYLE_MAP.get(xf.border.top_line_style, xlwt.Borders.NO_LINE)
    borders.bottom = LINE_STYLE_MAP.get(xf.border.bottom_line_style, xlwt.Borders.NO_LINE)
    borders.left_colour = xf.border.left_colour_index
    borders.right_colour = xf.border.right_colour_index
    borders.top_colour = xf.border.top_colour_index
    borders.bottom_colour = xf.border.bottom_colour_index
    style.borders = borders
    
    # 复制背景色
    pattern = xlwt.Pattern()
    pattern.pattern = xf.background.fill_pattern
    pattern.pattern_fore_colour = xf.background.pattern_colour_index
    pattern.pattern_back_colour = xf.background.background_colour_index
    style.pattern = pattern
    
    # 复制字体
    font = xlwt.Font()
    xf_font = book.font_list[xf.font_index]
    font.name = xf_font.name
    font.bold = xf_font.bold
    font.italic = xf_font.italic
    font.underline = xf_font.underline_type
    font.height = xf_font.height
    font.colour_index = xf_font.colour_index
    style.font = font
    
    # 复制对齐
    alignment = xlwt.Alignment()
    alignment.horz = xf.alignment.hor_align
    alignment.vert = xf.alignment.vert_align
    alignment.wrap = xf.alignment.text_wrapped
    style.alignment = alignment
    
    # 复制数字格式
    style.num_format_str = book.format_map[xf.format_key].format_str if xf.format_key in book.format_map else 'General'
    
    style_cache[cache_key] = style
    return style


def get_wrapped_cell_style(book, sheet, rowx, colx):
    """获取保留原格式且开启自动换行的单元格样式。"""
    style = pycopy.deepcopy(get_cell_style(book, sheet, rowx, colx))
    style.alignment.wrap = 1
    return style


def is_tuwen_shuoming_row(sheet, row_idx):
    """判断是否为图文说明行（第一处，含金额）"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val))
    row_text = ' '.join(row_vals)
    
    # 第一处图文说明：包含"图文说明:"且包含金额数字
    if '图文说明:' in row_text and re.search(r'\d+\.?\d*', row_text):
        # 排除第二处（下单日期附近的那行）
        if '下单日期' not in row_text and '交货期' not in row_text:
            return True
    return False


def is_tuwen_shuoming_row2(sheet, row_idx):
    """判断是否为第二处图文说明行（下单日期附近，不含金额）"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val))
    row_text = ' '.join(row_vals)
    
    if '图文说明:' in row_text and ('下单日期' in row_text or '交货期' in row_text):
        return True
    return False


def should_hide_row_digital(sheet, row_idx):
    """判断数字sheet的某行是否应该隐藏"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val))
    row_text = ' '.join(row_vals)

    # 不再隐藏第一处图文说明行（改为提取处理）
    # if '图文说明' in row_text:
    #     if re.search(r'\d+\.?\d*', row_text):
    #         return True

    # 隐藏包含"板材"或"五金"的行
    if '板材' in row_text or '五金' in row_text:
        return True

    # 隐藏包含"此单共"或"合计金额"的行
    if '此单共' in row_text or '合计金额' in row_text:
        return True

    # 隐藏"经销商确认"行
    if '经销商确认' in row_text:
        return True

    return False


def should_hide_row_hardware(sheet, row_idx):
    """判断五金sheet的某行是否应该隐藏"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val))
    row_text = ' '.join(row_vals)

    if '此单共' in row_text or '合计金额' in row_text:
        return True

    if '优惠' in row_text or '需支付' in row_text:
        return True

    return False


def should_hide_row_waigou(sheet, row_idx):
    """判断外购表/特殊产品表的某行是否应该隐藏"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val))
    row_text = ' '.join(row_vals)

    if '\u6570\u91cf\u603b\u8ba1' in row_text:
        return False

    # 隐藏合计/汇总/金额行
    if any(k in row_text for k in ['合计', '汇总', '总价格', '金额']):
        # 排除纯表头行（如"金额"作为列标题）
        # 如果整行很短（只有1-2个非空值）且包含这些词，更可能是合计行
        non_empty = [v for v in row_vals if str(v).strip()]
        if len(non_empty) <= 4:
            return True
    
    return False


def is_waigou_like_sheet(sheet_name):
    name = str(sheet_name)
    return '外购表' in name or '特殊产品表' in name


def find_and_replace_date(sheet, ws, book, order_date):
    """查找并替换下单日期（支持标签和日期在不同单元格的情况）"""
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            val = sheet.cell(row, col).value
            if val and isinstance(val, str):
                if '下单日期' in val or '接单日期' in val:
                    # 情况1：标签和日期在同一个单元格
                    new_val = re.sub(r'(下单日期[：:]\s*)\d{4}[\.\-]\d{1,2}[\.\-]\d{1,2}', r'\1' + order_date, val)
                    if new_val != val:
                        style = get_cell_style(book, sheet, row, col)
                        ws.write(row, col, new_val, style)
                    else:
                        # 情况2：标签和日期在相邻单元格
                        # 检查右侧或下方单元格是否有日期
                        for dc in [1, 2, 3, 4, 5]:  # 向右检查1-5列
                            next_col = col + dc
                            if next_col < sheet.ncols:
                                next_val = sheet.cell(row, next_col).value
                                if next_val and re.match(r'^\d{1,4}[\.\-]\d{1,2}([\.\-]\d{1,2})?$', str(next_val)):
                                    style = get_cell_style(book, sheet, row, next_col)
                                    ws.write(row, next_col, order_date, style)
                                    break
                            # 检查下方单元格
                            for dr in [1, 2]:
                                next_row = row + dr
                                if next_row < sheet.nrows:
                                    next_val = sheet.cell(next_row, col).value
                                    if next_val and re.match(r'^\d{1,4}[\.\-]\d{1,2}([\.\-]\d{1,2})?$', str(next_val)):
                                        style = get_cell_style(book, sheet, next_row, col)
                                        ws.write(next_row, col, order_date, style)
                                        break


def process_tuwen_shuoming(sheet, ws, book, row_idx):
    """处理图文说明行：保留"图文说明:"文本，清空金额列，同时把文本写到下面空白的图文说明行"""
    # 先找到下面空白的图文说明行（含"图文说明:"但不含金额）
    target_row = None
    for r in range(row_idx + 1, sheet.nrows):
        row_vals = []
        for c in range(sheet.ncols):
            v = sheet.cell(r, c).value
            if v:
                row_vals.append(str(v))
        row_text = ' '.join(row_vals)
        if '图文说明:' in row_text and not re.search(r'\d+\.?\d*', row_text):
            target_row = r
            break
    
    for col in range(sheet.ncols):
        val = sheet.cell(row_idx, col).value
        style = get_cell_style(book, sheet, row_idx, col)
        
        if val and isinstance(val, str) and '图文说明:' in val:
            # 保留"图文说明:"文本，去掉后面的金额
            new_val = '图文说明:'
            ws.write(row_idx, col, new_val, style)
            # 同时写到下面的空白图文说明行
            if target_row is not None:
                ws.write(target_row, col, new_val, style)
        elif val and (isinstance(val, (int, float)) or re.match(r'^\d+\.?\d*$', str(val))):
            # 清空金额数字
            ws.write(row_idx, col, '', style)
        else:
            # 保留其他内容
            ws.write(row_idx, col, val, style)


def should_hide_row_by_content(sheet, row_idx):
    """根据内容特征判断某行是否应该隐藏（纯内容判断，不依赖行号/模板）"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val).strip())
    row_text = ' '.join(row_vals)
    
    if not row_vals:
        return False
    
    # 1. 含金额的图文说明行（如"图文说明: 282.722 元"）
    if '图文说明:' in row_text:
        # 如果同时包含"下单日期"或"交货期"/"交货日期"，这是空白图文说明框，不隐藏
        if '下单日期' in row_text or '交货期' in row_text or '交货日期' in row_text:
            return False
        # 如果包含金额数字，隐藏
        if re.search(r'\d+\.?\d*', row_text):
            return True
    
    # 2. 板材汇总行（包含"板材"且有数字）
    if '板材' in row_text:
        # 检查是否有金额数字（排除纯"板材"标题）
        for v in row_vals:
            if re.search(r'\d+\.?\d*', v):
                return True
    
    # 3. 五金汇总行（包含"五金"且有数字）
    if '五金' in row_text:
        for v in row_vals:
            if re.search(r'\d+\.?\d*', v):
                return True
    
    # 4. 合计金额行
    if '此单共' in row_text or '合计金额' in row_text:
        return True
    
    # 5. 经销商确认行
    if '经销商确认' in row_text:
        return True
    
    # 6. 定金行（包含"定金"且有数字）
    if '定金' in row_text:
        for v in row_vals:
            if re.search(r'\d+\.?\d*', v):
                return True
    
    return False


def is_process_area_row(sheet, row_idx):
    """判断是否为加工流程区的行（应该保留显示）"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val).strip())
    row_text = ' '.join(row_vals)
    
    process_keywords = ['柜体加工', '门板加工', '电子锯', '雕刻', '打磨', '喷胶',
                        '小料开料', '封边', '排孔', '异形', '开槽', '试装', '包装',
                        '入库', '发货', '件数', '吸塑', '包覆', '开榫', '门芯加工',
                        '门芯吸塑', '拼装', '打包',
                        '设计', '开料', '制作', '贴皮', '木磨', '灰工',
                        '一次油磨', '二次油磨', '底漆', '擦色', '面漆']
    return any(kw in row_text for kw in process_keywords)


def is_zhibiaoren_row(sheet, row_idx):
    """判断是否为加工流程区下面的制表人/审单行（应该保留显示）"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val).strip())
    row_text = ' '.join(row_vals)
    
    if '制表人' not in row_text or '审单' not in row_text:
        return False
    
    # 检查前面8行内是否有加工流程区
    for rr in range(max(0, row_idx - 8), row_idx):
        if is_process_area_row(sheet, rr):
            return True
    return False


def is_blank_tuwen_box(sheet, row_idx):
    """判断是否为空白图文说明框（含下单日期/交货期，应该保留显示）"""
    row_vals = []
    for col_idx in range(sheet.ncols):
        val = sheet.cell(row_idx, col_idx).value
        if val:
            row_vals.append(str(val).strip())
    row_text = ' '.join(row_vals)
    
    return '图文说明:' in row_text and ('下单日期' in row_text or '交货期' in row_text or '交货日期' in row_text)


def is_blank_row(sheet, row_idx):
    """判断某行是否完全空白（所有单元格均无内容）"""
    for col_idx in range(sheet.ncols):
        if sheet.cell(row_idx, col_idx).value:
            return False
    return True


def get_row_signature(sheet, row_idx):
    """获取行的内容签名（用于匹配相同模式的行）"""
    vals = []
    for c in range(sheet.ncols):
        v = sheet.cell(row_idx, c).value
        if v:
            vals.append(str(v).strip())
    return tuple(vals)


def row_matches_template(sheet, row_idx, template_sig):
    """判断某行是否与模板签名匹配（支持内容模式匹配，不完全相同）"""
    sig = get_row_signature(sheet, row_idx)
    
    # 如果完全相同，直接匹配
    if sig == template_sig:
        return True
    
    # 如果模板是图文说明行（第一个元素是"图文说明:"）
    # 只匹配含金额的图文说明行，不匹配空白图文说明框
    if template_sig and template_sig[0] == '图文说明:':
        if sig and len(sig) >= 2 and sig[0] == '图文说明:':
            # 检查是否包含金额数字（排除空白图文说明框）
            row_text = ' '.join(sig)
            if re.search(r'\d+\.?\d*', row_text):
                return True
    
    # 如果模板是板材汇总行（包含"板材"）
    if template_sig and any('板材' in s for s in template_sig):
        if sig and any('板材' in s for s in sig):
            return True
    
    # 如果模板是合计金额行（包含"此单共"或"合计金额"）
    if template_sig and any('此单共' in s or '合计金额' in s for s in template_sig):
        if sig and any('此单共' in s or '合计金额' in s for s in sig):
            return True
    
    # 如果模板是经销商确认行
    if template_sig and any('经销商确认' in s for s in template_sig):
        if sig and any('经销商确认' in s for s in sig):
            return True
    
    # 如果模板是制表人行（包含"制表人"）
    # 但只匹配金额汇总区下面的制表人（后面没有柜体加工/门板加工的）
    if template_sig and any('制表人' in s for s in template_sig):
        if sig and any('制表人' in s for s in sig):
            # 检查这个制表人行后面是否有加工流程区
            has_process_after = False
            for rr in range(row_idx + 1, sheet.nrows):
                rv = []
                for cc in range(sheet.ncols):
                    vv = sheet.cell(rr, cc).value
                    if vv:
                        rv.append(str(vv))
                rt = ' '.join(rv)
                if '柜体加工' in rt or '门板加工' in rt:
                    has_process_after = True
                    break
            # 如果后面有加工流程区，说明这是加工流程区下面的制表人，不匹配
            if not has_process_after:
                return True
    
    return False


def find_matching_rows(sheet, template_rows, template_book=None, template_sheet_name=None):
    """找到与模板行内容模式匹配的所有行"""
    # 如果提供了模板book和sheet名，从模板获取签名；否则从当前sheet获取
    if template_book and template_sheet_name:
        template_sheet = template_book.sheet_by_name(template_sheet_name)
    else:
        template_sheet = sheet
    
    template_sigs = []
    for tr in template_rows:
        sig = get_row_signature(template_sheet, tr)
        template_sigs.append(sig)
    
    matched = []
    for r in range(sheet.nrows):
        for tsig in template_sigs:
            if row_matches_template(sheet, r, tsig):
                matched.append(r)
                break
    return matched


def find_tuwen_target_row(sheet, source_row):
    """找到图文说明内容应该写到的目标行（下面最近的空白图文说明行）"""
    for r in range(source_row + 1, sheet.nrows):
        row_vals = []
        for c in range(sheet.ncols):
            v = sheet.cell(r, c).value
            if v:
                row_vals.append(str(v))
        row_text = ' '.join(row_vals)
        # 目标行：含"图文说明:" 或 "下单日期" 或 "交货期" 或 "包装预计交货期"
        if '图文说明:' in row_text or '下单日期' in row_text or '交货期' in row_text:
            return r
    return None


def _get_date_source_sheet_name(book):
    """获取日期源 sheet 名称，优先级：1五/1五金 > 1"""
    for name in book.sheet_names():
        if name in ('1五', '1五金'):
            return name
    return '1' if '1' in book.sheet_names() else None


def _extract_dates_from_source(book):
    """从源文件提取统一日期值（优先从五金表读取，否则从Sheet '1'读取）"""
    order_date = None
    delivery_date = None
    
    date_source = _get_date_source_sheet_name(book)
    
    # 优先从日期源 sheet 读取
    if date_source:
        sheet = book.sheet_by_name(date_source)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if val and isinstance(val, str):
                    if '下单日期' in val and col + 1 < sheet.ncols:
                        next_val = sheet.cell(row, col + 1).value
                        if next_val and re.match(r'^\d{1,4}[\.\-]\d{1,2}([\.\-]\d{1,2})?$', str(next_val)):
                            order_date = str(next_val).strip()
                    elif ('预计交货期' in val or '预计交货日期' in val) and col + 1 < sheet.ncols:
                        next_val = sheet.cell(row, col + 1).value
                        if next_val and re.match(r'^\d{1,4}[\.\-]\d{1,2}([\.\-]\d{1,2})?$', str(next_val)):
                            delivery_date = str(next_val).strip()
    
    return order_date, delivery_date


def _clear_sheet1_dates(sheet, ws, book):
    """清空Sheet '1'的日期值（M列），保留标签"""
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            val = sheet.cell(row, col).value
            if val and isinstance(val, str):
                if '下单日期' in val or '预计交货期' in val or '预计交货日期' in val:
                    # 清空右侧M列(12)的值
                    if col + 1 < sheet.ncols:
                        next_col = col + 1
                        style = get_cell_style(book, sheet, row, next_col)
                        ws.write(row, next_col, '', style)


def _fill_dates_to_sheet(sheet, ws, book, order_date, delivery_date):
    """将统一日期值填入sheet的日期位置"""
    if not order_date and not delivery_date:
        return
    
    sheet_name = sheet.name
    
    # 判断sheet类型
    is_wu_sheet = '五' in sheet_name or '五金' in sheet_name
    is_number_sheet = sheet_name in [str(i) for i in range(1, 100)]
    
    if is_wu_sheet:
        # 五金sheet：日期在底部行
        # 找到包含 "下单日期" 和 "包装预计交货日期" 的行
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if val and isinstance(val, str):
                    if '下单日期' in val and order_date:
                        # 值在 D 列（索引 3），检查是否是合并单元格区域
                        # 找到值所在的列（通常是标签右侧第一个非空列或固定D列）
                        val_col = 3  # D列
                        if val_col < sheet.ncols:
                            style = get_cell_style(book, sheet, row, val_col)
                            ws.write(row, val_col, order_date, style)
                    elif ('预计交货日期' in val or '预计交货期' in val) and delivery_date:
                        val_col = 13  # N列
                        if val_col < sheet.ncols:
                            style = get_cell_style(book, sheet, row, val_col)
                            ws.write(row, val_col, delivery_date, style)
    
    elif is_number_sheet:
        # 数字sheet（报价单）：日期在 L/M 列
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if val and isinstance(val, str):
                    if '下单日期' in val and order_date:
                        # 值在 M 列（索引 12，标签右侧）
                        val_col = col + 1
                        if val_col < sheet.ncols:
                            style = get_cell_style(book, sheet, row, val_col)
                            ws.write(row, val_col, order_date, style)
                    elif ('预计交货期' in val or '预计交货日期' in val) and delivery_date:
                        val_col = col + 1
                        if val_col < sheet.ncols:
                            style = get_cell_style(book, sheet, row, val_col)
                            ws.write(row, val_col, delivery_date, style)
    
    else:
        # 其他sheet（特殊产品表、实木、汇总等）：按通用逻辑处理
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if val and isinstance(val, str):
                    if '下单日期' in val and order_date:
                        for dc in range(1, 6):
                            if col + dc < sheet.ncols:
                                style = get_cell_style(book, sheet, row, col + dc)
                                ws.write(row, col + dc, order_date, style)
                                break
                    elif ('预计交货日期' in val or '预计交货期' in val) and delivery_date:
                        for dc in range(1, 6):
                            if col + dc < sheet.ncols:
                                style = get_cell_style(book, sheet, row, col + dc)
                                ws.write(row, col + dc, delivery_date, style)
                                break


def generate_factory_version(source_path, output_path, order_date=None):
    """生成工厂版：隐藏价格列和行，保留原文件格式（用xlutils复制），日期替换"""
    import xlrd
    from xlutils.copy import copy
    import xlwt
    
    # 清空样式缓存
    global style_cache
    style_cache = {}
    
    book = xlrd.open_workbook(source_path, formatting_info=True)
    wb = copy(book)
    
    # 确定日期源 sheet，优先级：1五/1五金 > 1
    date_source = _get_date_source_sheet_name(book)
    
    # 从日期源 sheet 扫描日期单元格位置，用于其他 sheet 写公式引用
    date_refs = {}
    wu_layout_template = None
    
    def _is_date_value(cell_val, book, s, r, c):
        """判断单元格值是否为日期（支持字符串日期和Excel日期数字）"""
        if not cell_val and cell_val != 0:
            return False
        if isinstance(cell_val, str):
            stripped = cell_val.strip()
            if re.match(r'^\d{1,4}[\.\-]\d{1,2}([\.\-]\d{1,2})?$', stripped):
                return True
        # xlrd 日期类型返回 float (xldate)
        if isinstance(cell_val, (int, float)) and 0 < cell_val < 100000:
            try:
                if s.cell(r, c).ctype == xlrd.XL_CELL_DATE:
                    return True
            except Exception:
                pass
            # 备用：数字格式也视为日期
            xf_index = s.cell_xf_index(r, c)
            xf = book.xf_list[xf_index]
            fmt = book.format_map.get(xf.format_key) if hasattr(book, 'format_map') else None
            if fmt and ('date' in fmt.format_str.lower() or 'yy' in fmt.format_str.lower() or 'mm' in fmt.format_str.lower() or 'dd' in fmt.format_str.lower()):
                return True
        return False

    def _find_date_cell(s_src, row, col):
        """找到日期标签右侧真正含日期值的单元格引用。"""
        for dc in range(1, 6):
            nc = col + dc
            if nc >= s_src.ncols:
                break
            nv = s_src.cell(row, nc).value
            if _is_date_value(nv, book, s_src, row, nc):
                return Utils.rowcol_to_cell(row, nc), nv
        return None, None
    
    def _find_date_target_col(sheet, row, col, is_wu_sheet):
        """找到日期标签对应的写入列；五金表优先按合并块右侧单元格写入。"""
        if is_wu_sheet:
            for rlo, rhi, clo, chi in getattr(sheet, "merged_cells", []):
                if rlo <= row < rhi and clo <= col < chi:
                    if chi < sheet.ncols:
                        return chi
                    return None

        target_col = None
        for dc in range(1, 10 if is_wu_sheet else 6):
            next_col = col + dc
            if next_col >= sheet.ncols:
                break
            next_val = sheet.cell(row, next_col).value
            if next_val:
                target_col = next_col
                break
            elif target_col is None:
                target_col = next_col
        return target_col
    
    def _scan_date_info(s_src, refs_dict, vals_dict, is_wu=False):
        """从指定 sheet 扫描日期引用位置和实际值。"""
        for r in range(s_src.nrows):
            for c in range(s_src.ncols):
                v = s_src.cell(r, c).value
                if v and isinstance(v, str):
                    ref_key = None
                    if '接单日期' in v:
                        ref_key = 'receipt_date'
                    elif '下单日期' in v:
                        ref_key = 'order_date'
                    elif '预计交货日期' in v or '预计交货期' in v or '包装预计交货' in v:
                        ref_key = 'delivery_date'
                    
                    if ref_key and ref_key not in refs_dict:
                        ref, val = _find_date_cell(s_src, r, c)
                        if ref:
                            refs_dict[ref_key] = ref
                            if val is not None:
                                vals_dict[ref_key] = val
                        else:
                            # 使用 _find_date_target_col 找到正确目标位置（含合并单元格处理）
                            tc = _find_date_target_col(s_src, r, c, is_wu)
                            if tc is not None:
                                refs_dict[ref_key] = Utils.rowcol_to_cell(r, tc)
    
    date_values = {}  # 实际日期值 {ref_key: value}
    
    # 1. 优先从五金表扫描日期位置
    if date_source and date_source != '1':
        s_src = book.sheet_by_name(date_source)
        _scan_date_info(s_src, date_refs, date_values, is_wu=True)
    
    # 2. 从普通表 1 补充缺失的日期值（但不改变 date_source）
    if '1' in book.sheet_names():
        s1 = book.sheet_by_name('1')
        _scan_date_info(s1, {}, date_values, is_wu=False)
    
    # 3. 如果根本没有五金表，从普通表 1 扫描位置
    if date_source == '1' and '1' in book.sheet_names():
        s_src = book.sheet_by_name('1')
        _scan_date_info(s_src, date_refs, date_values, is_wu=False)
    
    def _write_formula_if_in_bounds(ws_obj, style_sheet, row_idx, col_idx, formula_text):
        """写公式（带边界保护和原样式保留）"""
        if row_idx < 0 or col_idx < 0:
            return
        if row_idx >= style_sheet.nrows or col_idx >= style_sheet.ncols:
            return
        style = get_cell_style(book, style_sheet, row_idx, col_idx)
        ws_obj.write(row_idx, col_idx, xlwt.Formula(formula_text), style)

    def _get_output_style(sheet, row, col, wrap_text=False):
        if wrap_text:
            return get_wrapped_cell_style(book, sheet, row, col)
        return get_cell_style(book, sheet, row, col)

    def _find_row_containing(sheet, keyword, start_row=0):
        """查找包含指定文本的第一行。"""
        for row in range(start_row, sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if val and keyword in str(val):
                    return row
        return None

    def _load_wu_layout_template():
        """读取五金表行高列宽模板。"""
        for template_path in WU_LAYOUT_TEMPLATE_PATHS:
            if not os.path.exists(template_path):
                continue
            try:
                template_book = xlrd.open_workbook(template_path, formatting_info=True)
                template_sheet = template_book.sheet_by_name(WU_LAYOUT_TEMPLATE_SHEET)
            except Exception:
                continue

            col_widths = {}
            for col, col_info in template_sheet.colinfo_map.items():
                col_widths[col] = col_info.width

            row_heights = {}
            for row, row_info in template_sheet.rowinfo_map.items():
                row_heights[row] = row_info.height

            table_start = _find_row_containing(template_sheet, '序号')
            table_end = None
            if table_start is not None:
                table_end = _find_row_containing(template_sheet, '数量总计', table_start)

            return {
                'col_widths': col_widths,
                'row_heights': row_heights,
                'table_start': table_start,
                'table_end': table_end,
            }
        return None

    def _find_normal_bottom_start(sheet, sheet_start):
        """找到普通表数据区结束后的第一行（底部区域开始）。"""
        for row in range(sheet_start + 1, sheet.nrows):
            row_vals = []
            for col in range(sheet.ncols):
                v = sheet.cell(row, col).value
                if v:
                    row_vals.append(str(v).strip())
            row_text = ' '.join(row_vals)
            if '柜体加工' in row_text or '门板加工' in row_text:
                return row
            if '包装预计交货' in row_text:
                return row
            if '图文说明:' in row_text and ('下单日期' in row_text or '交货期' in row_text):
                return row
        return sheet.nrows

    def _apply_normal_standard_layout(ws_obj, sheet):
        """按标准常量套用普通表尺寸，不重写任何单元格内容或公式。"""
        for col, width in NORMAL_STANDARD_COL_WIDTHS.items():
            if col < sheet.ncols:
                ws_obj.col(col).width = width

        sheet_start = _find_row_containing(sheet, '序号')
        bottom_start = _find_normal_bottom_start(sheet, sheet_start) if sheet_start is not None else None
        row_heights = NORMAL_STANDARD_ROW_HEIGHTS

        if sheet_start is None or bottom_start is None:
            for row, height in row_heights.items():
                if row < sheet.nrows:
                    ws_obj.row(row).height = height
            return

        # 表头及表头前按标准行号
        for row in range(min(sheet_start, sheet.nrows)):
            if row in row_heights:
                ws_obj.row(row).height = row_heights[row]

        # 表头行套 560，数据区从下一行开始套 440
        ws_obj.row(sheet_start).height = NORMAL_STANDARD_ROW_HEIGHTS[6]
        for row in range(sheet_start + 1, min(bottom_start, sheet.nrows)):
            ws_obj.row(row).height = NORMAL_STANDARD_ROW_HEIGHTS[7]

        # 底部区域按标准表"数据结束后"的相对行高套用
        std_data_end = 11  # 标准表最后一个数据行
        for offset, template_row in enumerate(range(std_data_end + 1, max(row_heights.keys(), default=std_data_end) + 1)):
            sheet_row = bottom_start + offset
            if sheet_row >= sheet.nrows:
                break
            if template_row in row_heights:
                ws_obj.row(sheet_row).height = row_heights[template_row]

    def _apply_wu_layout_template(ws_obj, sheet):
        """按标准常量套用五金表尺寸，不重写任何单元格内容或公式。"""
        for col, width in WU_STANDARD_COL_WIDTHS.items():
            if col < sheet.ncols:
                ws_obj.col(col).width = width

        sheet_start = _find_row_containing(sheet, '序号')
        sheet_end = _find_row_containing(sheet, '数量总计', sheet_start) if sheet_start is not None else None
        row_heights = WU_STANDARD_ROW_HEIGHTS

        if sheet_start is None or sheet_end is None:
            for row, height in row_heights.items():
                if row < sheet.nrows:
                    ws_obj.row(row).height = height
            return

        # 表头及表头前按标准行号
        for row in range(min(sheet_start, sheet.nrows)):
            if row in row_heights:
                ws_obj.row(row).height = row_heights[row]

        # 数据区及数量总计行每行用 559
        for row in range(sheet_start, min(sheet_end + 1, sheet.nrows)):
            ws_obj.row(row).height = 559

        # 底部区域按标准表"数量总计后"的相对行高套用
        std_end = 23  # 标准表数量总计行
        for offset, template_row in enumerate(range(std_end + 1, max(row_heights.keys(), default=std_end) + 1)):
            sheet_row = sheet_end + 1 + offset
            if sheet_row >= sheet.nrows:
                break
            if template_row in row_heights:
                ws_obj.row(sheet_row).height = row_heights[template_row]

    def _apply_wu_print_settings(ws_obj):
        """五金表打印设置：A4竖向，所有列适配到一页宽。"""
        ws_obj.paper_size_code = 9
        ws_obj.portrait = 1
        ws_obj.fit_num_pages = 1
        ws_obj.fit_width_to_pages = 1
        ws_obj.fit_height_to_pages = 0
        ws_obj.print_scaling = 100
        ws_obj.left_margin = 0.25
        ws_obj.right_margin = 0.25
        ws_obj.top_margin = 0.35
        ws_obj.bottom_margin = 0.35
        ws_obj.header_margin = 0.1
        ws_obj.footer_margin = 0.1

    def _load_normal_layout_template():
        """读取普通表行高列宽模板。"""
        for template_path in NORMAL_LAYOUT_TEMPLATE_PATHS:
            if not os.path.exists(template_path):
                continue
            try:
                template_book = xlrd.open_workbook(template_path, formatting_info=True)
                template_sheet = template_book.sheet_by_name(NORMAL_LAYOUT_TEMPLATE_SHEET)
            except Exception:
                continue

            col_widths = {}
            for col, col_info in template_sheet.colinfo_map.items():
                col_widths[col] = col_info.width

            row_heights = {}
            for row, row_info in template_sheet.rowinfo_map.items():
                row_heights[row] = row_info.height

            table_start = _find_row_containing(template_sheet, '序号')
            table_end = None
            if table_start is not None:
                # 模板中第一个超高行 (>2000) 之前视为数据区结尾
                for r in range(table_start + 1, template_sheet.nrows):
                    if r in template_sheet.rowinfo_map and template_sheet.rowinfo_map[r].height > 2000:
                        table_end = r - 1
                        break
                if table_end is None:
                    table_end = template_sheet.nrows - 1

            return {
                'col_widths': col_widths,
                'row_heights': row_heights,
                'table_start': table_start,
                'table_end': table_end,
            }
        return None

    def _apply_normal_layout_template(ws_obj, sheet):
        """按标准常量套用普通表尺寸，不重写任何单元格内容或公式。"""
        _apply_normal_standard_layout(ws_obj, sheet)

    wu_layout_template = _load_wu_layout_template()

    # 处理每个sheet：先补全所有单元格格式（xlutils.copy对某些xls格式复制不完整）
    for sheet_idx, sheet_name in enumerate(book.sheet_names()):
        sheet = book.sheet_by_name(sheet_name)
        ws = wb.get_sheet(sheet_idx)
        is_number_sheet = sheet_name in [str(i) for i in range(1, 100)]
        is_wu_sheet = '五' in sheet_name or '五金' in sheet_name
        
        # 找到最后有内容的行，避免尾部空白行保留格式导致空白页
        last_content_row = -1
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                if sheet.cell(r, c).value:
                    last_content_row = r
                    break
        
        # 显式复制所有单元格的边框和样式，确保不丢失
        # 尾部空白行（内容之后的行）不写格式，避免产生空白打印页
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if row > last_content_row and not val:
                    ws.write(row, col, '')
                else:
                    style = _get_output_style(sheet, row, col, is_wu_sheet)
                    ws.write(row, col, val, style)
        
        # 1. 替换日期：普通表和五金表按不同字段匹配，找到标签后在右侧写入公式引用 Sheet "1"
        if order_date or date_refs:
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    val = sheet.cell(row, col).value
                    if not val or not isinstance(val, str):
                        continue
                    
                    # 根据表类型选择匹配的字段
                    ref_key = None
                    is_delivery = False
                    if is_wu_sheet:
                        # 五金表：找 "下单日期" / "接单日期"、"包装预计交货日期"
                        if '接单日期' in val:
                            ref_key = 'receipt_date'
                        elif '下单日期' in val:
                            ref_key = 'order_date'
                        elif '包装预计交货日期' in val:
                            is_delivery = True
                    else:
                        # 普通表：找 "下单日期" / "接单日期"、"包装预计交货期" / "包装预计交货日期"
                        if '接单日期' in val:
                            ref_key = 'receipt_date'
                        elif '下单日期' in val:
                            ref_key = 'order_date'
                        elif '包装预计交货期' in val or '包装预计交货日期' in val:
                            is_delivery = True
                    
                    if ref_key is None and not is_delivery:
                        continue
                    
                    # 找到标签对应的日期单元格；五金表底部通常是合并标签 + 右侧合并值。
                    # 外购表复用五金表的合并单元格右侧规则
                    is_wu_like = is_wu_sheet or is_waigou_like_sheet(sheet_name)
                    target_col = _find_date_target_col(sheet, row, col, is_wu_like)
                    
                    if target_col is not None:
                        style = _get_output_style(sheet, row, target_col, is_wu_sheet)
                        if sheet_name == date_source:
                            # 日期源 sheet 写入实际值：优先传入参数，其次从其他 sheet 提取的日期值
                            if order_date:
                                ws.write(row, target_col, order_date, style)
                            elif ref_key and ref_key in date_values:
                                ws.write(row, target_col, date_values[ref_key], style)
                            elif is_delivery and 'delivery_date' in date_values:
                                ws.write(row, target_col, date_values['delivery_date'], style)
                            # 如果日期源自身已有值，保留（步骤2已写入）
                        else:
                            # 其他 sheet 统一写公式引用日期源
                            if ref_key and ref_key in date_refs:
                                ws.write(row, target_col, xlwt.Formula(f"'{date_source}'!" + date_refs[ref_key]), style)
                            elif ref_key and 'order_date' in date_refs:
                                ws.write(row, target_col, xlwt.Formula(f"'{date_source}'!" + date_refs['order_date']), style)
                            elif is_delivery and 'delivery_date' in date_refs:
                                ws.write(row, target_col, xlwt.Formula(f"'{date_source}'!" + date_refs['delivery_date']), style)
        


        # 2. 隐藏价格列
        hidden_cols = set()
        if is_number_sheet:
            hidden_cols = {7, 8, 9, 10}  # H, I, J, K
        elif is_wu_sheet:
            hidden_cols = {10, 16, 17}  # K, Q, R
        elif is_waigou_like_sheet(sheet_name):
            hidden_cols = {10, 11, 12}  # K, L, M
            # 检测第二种外购表结构（单价在 Col 9，表头中有"平方"在 Col 10）
            for r in range(min(10, sheet.nrows)):
                for c in range(sheet.ncols):
                    v = sheet.cell(r, c).value
                    if v and isinstance(v, str) and '平方' in v:
                        hidden_cols.add(9)  # J 列（单价）
                        break
                if 9 in hidden_cols:
                    break
        elif sheet_name == '实木':
            hidden_cols = {10, 11}  # K, L
        elif sheet_name == '汇总':
            hidden_cols = {1, 2, 3, 4}  # B, C, D, E
        
        for col_idx in hidden_cols:
            if col_idx < sheet.ncols:
                ws.col(col_idx).hidden = 1

        if is_wu_sheet:
            _apply_wu_layout_template(ws, sheet)
            _apply_wu_print_settings(ws)
        elif is_waigou_like_sheet(sheet_name):
            # 外购表固定列宽：G列=21, J列=14
            if sheet.ncols > 6:
                ws.col(6).width = 21 * 256
            if sheet.ncols > 9:
                ws.col(9).width = 14 * 256
        
        # 3. 隐藏右侧所有无内容的空列，并给最后一列可见列补右边框
        last_visible_col = 0
        for col in range(sheet.ncols):
            if col in hidden_cols:
                continue
            has_content = False
            for row in range(sheet.nrows):
                if sheet.cell(row, col).value:
                    has_content = True
                    break
            if has_content:
                last_visible_col = col

        # 隐藏空白列
        for col in range(last_visible_col + 1, sheet.ncols):
            ws.col(col).hidden = 1
        
        # 给最后一列可见列的每个单元格补右边框（复制左侧相邻列的右边框样式）
        if last_visible_col > 0 and last_visible_col < sheet.ncols - 1:
            for row in range(sheet.nrows):
                # 获取原始该单元格的样式
                xf_index = sheet.cell_xf_index(row, last_visible_col)
                xf = book.xf_list[xf_index]
                
                # 如果该单元格本身没有右边框，从隐藏列的左边框复制
                if xf.border.right_line_style == 0 and last_visible_col + 1 in hidden_cols:
                    next_xf_index = sheet.cell_xf_index(row, last_visible_col + 1)
                    next_xf = book.xf_list[next_xf_index]
                    if next_xf.border.left_line_style > 0:
                        style = _get_output_style(sheet, row, last_visible_col, is_wu_sheet)
                        # 只修改右边框为相邻隐藏列的左边框
                        style.borders.right = LINE_STYLE_MAP.get(next_xf.border.left_line_style, xlwt.Borders.THIN)
                        style.borders.right_colour = next_xf.border.left_colour_index
                        val = sheet.cell(row, last_visible_col).value
                        ws.write(row, last_visible_col, val, style)
                        continue
                
                # 如果单元格本身有右边框但颜色不对，保持原样（上面已经通过xlutils.copy复制了）
                # 这里只在确实需要补边框时才写入

        
        # 3. 隐藏价格行（xlwt方式：设置行高为0）
        if is_number_sheet or sheet_name == '实木':
            for row in range(sheet.nrows):
                # 实木：优先继承源文件已有的隐藏行状态
                if sheet_name == '实木':
                    if row in sheet.rowinfo_map and sheet.rowinfo_map[row].hidden:
                        ws.row(row).height = 0
                        ws.row(row).hidden = True
                        continue
                # 加工流程区、制表人/审单行、空白图文说明框：强制显示（xlutils.copy可能复制了原始文件的隐藏状态）
                if is_process_area_row(sheet, row) or is_zhibiaoren_row(sheet, row) or is_blank_tuwen_box(sheet, row):
                    ws.row(row).hidden = False
                    continue
                # 实木底部日期/生产单号行强制显示
                if sheet_name == '实木':
                    row_vals = []
                    for col_idx in range(sheet.ncols):
                        val = sheet.cell(row, col_idx).value
                        if val:
                            row_vals.append(str(val).strip())
                    row_text = ' '.join(row_vals)
                    if any(k in row_text for k in ['下单日期', '预计交货日期', '预计交货期', '生产单号']):
                        ws.row(row).hidden = False
                        continue
                if should_hide_row_by_content(sheet, row):
                    ws.row(row).height = 0
                    ws.row(row).hidden = True
        elif is_wu_sheet:
            for row in range(sheet.nrows):
                if should_hide_row_hardware(sheet, row):
                    ws.row(row).height = 0
                    ws.row(row).hidden = True
            # 动态隐藏"数量总计"下面紧跟的空白行
            total_qty_row = None
            for r in range(sheet.nrows):
                for c in range(sheet.ncols):
                    val = sheet.cell(r, c).value
                    if val and '数量总计' in str(val):
                        total_qty_row = r
                        break
                if total_qty_row is not None:
                    break
            if total_qty_row is not None:
                for r in range(total_qty_row + 1, sheet.nrows):
                    if is_blank_row(sheet, r):
                        ws.row(r).height = 0
                        ws.row(r).hidden = True
                    elif should_hide_row_hardware(sheet, r):
                        # 已被外层循环隐藏，继续往下扫空白行
                        continue
                    else:
                        break
        elif is_waigou_like_sheet(sheet_name):
            for row in range(sheet.nrows):
                row_text = ' '.join(
                    str(sheet.cell(row, col).value)
                    for col in range(sheet.ncols)
                    if sheet.cell(row, col).value
                )
                if '\u6570\u91cf\u603b\u8ba1' in row_text:
                    ws.row(row).hidden = False
                    continue
                if should_hide_row_waigou(sheet, row):
                    ws.row(row).height = 0
                    ws.row(row).hidden = True
    
    wb.save(output_path)
    return output_path


if __name__ == '__main__':
    import sys
    source = sys.argv[1] if len(sys.argv) > 1 else "/mnt/d/Edge下载/bom-server-1.3.0/B2604-4098马斌星_周周_.xls"
    output = sys.argv[2] if len(sys.argv) > 2 else "/tmp/test_factory_output_v2.xls"
    order_date = sys.argv[3] if len(sys.argv) > 3 else None
    generate_factory_version(source, output, order_date)
