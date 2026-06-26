import base64
import json
import urllib.request
from pathlib import Path

BASE = Path(r"C:/Users/Administrator/Downloads")
FILES = [
    ("综合查询5月.xlsx", "zhcx"),
    ("联思系统5月.xlsx", "liansi"),
    ("奢匠下单统计表5月.xlsx", "shejiang"),
    ("2026年5月销售部业绩核对表.xlsx", "template"),
]

payload_files = []
for fname, _ in FILES:
    fpath = BASE / fname
    with open(fpath, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    payload_files.append({"file_content": b64, "filename": fname})

payload = json.dumps({"files": payload_files}, ensure_ascii=False).encode("utf-8")

req = urllib.request.Request(
    "http://localhost:8003/process",
    data=payload,
    headers={"Content-Type": "application/json; charset=utf-8"},
)

with urllib.request.urlopen(req, timeout=300) as resp:
    result = json.loads(resp.read().decode("utf-8"))

print(json.dumps(result, ensure_ascii=False, indent=2))

if result.get("success"):
    for item in result.get("output_files", []):
        content = item.get("file_content")
        if content:
            out_path = Path("test_output.xlsx")
            out_path.write_bytes(base64.b64decode(content))
            print(f"已保存输出文件: {out_path.resolve()} ({out_path.stat().st_size} bytes)")
