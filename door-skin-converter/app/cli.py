from __future__ import annotations

import argparse
from pathlib import Path

from app.converter import ConversionError, convert_excel_to_csv


def main() -> int:
    parser = argparse.ArgumentParser(description="将门扇 Excel 转换为下单 CSV")
    parser.add_argument("input", type=Path, help="源文件（.xls 或 .xlsx）")
    parser.add_argument("-o", "--output", type=Path, help="输出 CSV 路径")
    args = parser.parse_args()

    output = args.output or args.input.with_name(f"{args.input.stem}_转换.csv")
    try:
        csv_bytes, stats = convert_excel_to_csv(args.input.read_bytes(), args.input.name)
    except (OSError, ConversionError, RuntimeError) as exc:
        parser.error(str(exc))

    output.write_bytes(csv_bytes)
    print(
        f"转换完成：源数据 {stats.source_rows} 行，输出 {stats.output_rows} 行，"
        f"数量合计 {stats.quantity_sum:g}；文件：{output}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

