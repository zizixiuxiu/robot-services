# -*- coding: utf-8 -*-
"""扫描 workshop-order 已保存的下车间文件，找出泄露价格（优惠/需支付/合计金额等横幅）的 xlsx。

用法：
    python scan_price_leak.py            # 只扫描并列出命中文件
    python scan_price_leak.py --delete   # 扫描并删除命中文件

只扫描 data/output 目录（服务生成的下车间成品存档）。
ishare-archive 下的"原文件"是输入原件，天然带价格，不在扫描范围内。
"""
import io
import os
import re
import glob
import sys
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "output")
PRICE_RE = re.compile(
    r"合计金额[:：]\s*\d|折后[:：]\s*\d|\d+折优惠|需支付|不打折|优惠[:：]\s*\d|优惠金额[:：]?\s*\d"
)


def scan():
    files = [p for p in glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True)]
    hits = []
    for path in files:
        try:
            with zipfile.ZipFile(path) as z:
                found = set()
                for name in z.namelist():
                    if name.endswith(".xml") and ("sheet" in name or "sharedStrings" in name):
                        data = z.read(name).decode("utf-8", errors="ignore")
                        for m in PRICE_RE.finditer(data):
                            s = max(0, m.start() - 40)
                            snippet = re.sub(r"<[^>]+>", "", data[s : m.end() + 40])[:80]
                            found.add(snippet)
                            if len(found) >= 2:
                                break
                    if len(found) >= 2:
                        break
            if found:
                hits.append((path, sorted(found)))
        except Exception as e:
            print(f"[读取失败] {path}: {e}")
    return hits


def main():
    delete = "--delete" in sys.argv
    hits = scan()
    print(f"扫描目录: {ROOT}")
    print(f"命中带价格文件: {len(hits)}")
    for path, found in hits:
        print(f"\n{path}")
        for f in found:
            print(f"    {f}")
        if delete:
            os.remove(path)
            print("    -> 已删除")
    if hits and not delete:
        print("\n（加 --delete 参数可删除以上文件）")


if __name__ == "__main__":
    main()
