#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV 板件转换核心逻辑（替代原 PowerShell 脚本）
输入：原始板件 CSV
输出：模板 CSV
"""

import csv
import re
import os
import sys
from pathlib import Path


OUTPUT_COLUMNS = [
    "合同自编号",
    "成型长",
    "成型宽",
    "成型高",
    "部件名称",
    "贴面纹路",
    "正面孔程序编码",
    "反面孔程序编码",
    "板件编码",
    "材料花色",
    "工艺编码",
    "开门方向",
    "备注1（生产信息）",
    "封边",
    "数量",
    "开料",
    "完工",
    "品牌",
]


def _extract_number(text):
    """从文本中提取第一个数字（支持负号和小数）"""
    if text is None:
        return ""
    text = str(text).strip()
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    return m.group(0) if m else ""


def _subtract_number(value, subtract_by=2):
    """提取数字并减去指定值，保留合适格式"""
    num_text = _extract_number(value)
    if num_text == "":
        return ""
    try:
        result = float(num_text) - subtract_by
        if result == int(result):
            return str(int(result))
        return f"{result:.3f}".rstrip("0").rstrip(".")
    except ValueError:
        return ""


def _replace_material_thickness(material, thickness):
    """把材料花色里的厚度替换为成型高，或前置厚度"""
    if not material or not thickness:
        return material or ""
    material = str(material).strip()
    thickness = str(thickness).strip()

    # 如果已包含 "数字mm"，替换该数字
    m = re.search(r"\d+(?:\.\d+)?(?=\s*mm)", material, re.IGNORECASE)
    if m:
        return material[:m.start()] + thickness + material[m.end():]

    # 否则前置厚度
    return f"{thickness}mm{material}"


def convert_csv(input_path, output_path, input_encoding="utf-8-sig", output_encoding="utf-8-sig"):
    """转换单个 CSV 文件"""
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_path}")

    # 尝试编码
    content = None
    for enc in (input_encoding, "utf-8", "gbk", "gb2312"):
        try:
            with open(input_path, "r", encoding=enc, newline="") as f:
                content = f.read()
            break
        except UnicodeDecodeError:
            continue
    if content is None:
        raise ValueError(f"无法解码输入文件: {input_path}")

    reader = csv.DictReader(content.splitlines())
    rows = list(reader)

    converted = []
    for row in rows:
        height = _extract_number(row.get("成型高", ""))
        converted.append({
            "合同自编号": row.get("合同自编号", ""),
            "成型长": row.get("成型长", ""),
            "成型宽": row.get("成型宽", ""),
            "成型高": row.get("成型高", ""),
            "部件名称": row.get("部件名称", ""),
            "贴面纹路": "1",
            "正面孔程序编码": row.get("正面孔程序编码", ""),
            "反面孔程序编码": row.get("反面孔程序编码", ""),
            "板件编码": row.get("板件编码", ""),
            "材料花色": _replace_material_thickness(row.get("材料花色", ""), height),
            "工艺编码": row.get("工艺编码", ""),
            "开门方向": row.get("开门方向", ""),
            "备注1（生产信息）": row.get("备注1（生产信息）", ""),
            "封边": "1-1-1-1-1-1-1-1-1-1",
            "数量": "1",
            "开料": _subtract_number(row.get("成型长", ""), 2),
            "完工": _subtract_number(row.get("成型宽", ""), 2),
            "品牌": "",
        })

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding=output_encoding, newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(converted)

    return {"success": True, "rows": len(converted), "output_path": str(output_path)}


def main():
    if len(sys.argv) < 3:
        print("用法: python convert_board_csv.py <输入CSV> <输出CSV>")
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    result = convert_csv(input_path, output_path)
    print(f"转换完成：{result['output_path']}，共 {result['rows']} 行")


if __name__ == "__main__":
    main()
