"""一条命令跑完「生成 + 改动审计」。

用法
----
    python audit_run.py 原文件.xls                 # 生成并打印改动清单
    python audit_run.py 原文件.xls --json out.json  # 同时导出明细
    python audit_run.py 原文件.xls --discount 0.85  # 换折扣系数

做的事
------
1. 把原始 .xls 转成 .xlsx（LibreOffice），这份就是「原表」基线；
2. 用同一份原文件跑 transform() 得到结果；
3. 逐格比对，打印每处改动及理由，最后断言「其余一字未改」。

注意
----
必须用同一份原文件同时产生基线和结果，否则比对的是两个不同版本的差异，
会凭空多出一堆假改动（实测踩过：拿 9/28 的基线比 11/5 的结果，多报了 5 处产品名被改）。
"""
from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import change_audit
import make_change_notice
import make_workshop_order as mwo

order_hint = make_change_notice.order_hint


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="生成下车间文件并审计每处改动")
    ap.add_argument("input", help="原始报价单（.xls 或 .xlsx）")
    ap.add_argument("--discount", type=float, default=mwo.DEFAULT_DISCOUNT)
    ap.add_argument("--order-type", default="auto", choices=sorted(mwo.ORDER_TYPES))
    ap.add_argument("--json", help="把改动明细写到这个 JSON")
    ap.add_argument("--limit", type=int, default=12, help="每张表最多列几条明细")
    ap.add_argument("--keep", help="保留基线与结果的目录（默认用临时目录，跑完删除）")
    args = ap.parse_args(argv)

    src = Path(args.input)
    if not src.exists():
        print(f"文件不存在：{src}")
        return 2

    with tempfile.TemporaryDirectory(prefix="audit_run_") as tmp:
        work = Path(args.keep) if args.keep else Path(tmp)
        work.mkdir(parents=True, exist_ok=True)
        base = work / "base.xlsx"
        gen = work / "gen.xlsx"

        if src.suffix.lower() == ".xls":
            print(f"[1/3] 转换 {src.name} → base.xlsx（LibreOffice）", flush=True)
            mwo.convert_xls_to_xlsx(src, base)
        else:
            print(f"[1/3] 直接以 {src.name} 作为基线", flush=True)
            base = src

        print("[2/3] 生成下车间文件", flush=True)
        stats = mwo.transform(src, gen, args.discount, args.order_type)
        print("      单价折扣 {} 处｜工分补填 {} 处｜木箱行清空 {} 处｜备注清理 {} 处".format(
            stats.get("discounted_prices", 0),
            stats.get("filled_workpoints", 0),
            stats.get("cleared_wood_boxes", 0),
            stats.get("updated_page_notes", 0),
        ), flush=True)

        print("[3/3] 逐格比对\n", flush=True)
        reports = change_audit.audit(base, gen)
        print(change_audit.render(reports, args.limit))

        payload = change_audit.to_payload(reports)
        items = [c for r in payload for c in r["cell_changes"]]
        # 每次生成都附一段可以直接发群的话，省得再人工组织语言
        print("\n" + "=" * 66)
        print("下面这段可以直接发到群里：")
        print("=" * 66 + "\n")
        print(make_change_notice.build_text(
            items, make_change_notice.summarize(items, payload), order_hint(src.stem)))

        if args.json:
            change_audit.write_json(reports, args.json)
            print(f"\n已保存明细：{args.json}")

        if args.keep:
            print(f"基线：{base}\n结果：{gen}")

        return 0 if not any(r.review() for r in reports) else 1


if __name__ == "__main__":
    sys.exit(main())
