#!/usr/bin/env python3
"""测试五金汇总 8001 服务的 /process 接口"""
import os
import sys
import json
import base64
import time
import urllib.request

SERVICE_URL = os.getenv("SERVICE_URL", "http://127.0.0.1:8001")


def test_health():
    req = urllib.request.Request(f"{SERVICE_URL}/health", method="GET")
    with urllib.request.urlopen(req, timeout=10) as resp:
        body = resp.read().decode("utf-8")
        print("health:", body)
        return resp.status == 200


def test_process(input_path: str):
    if not os.path.exists(input_path):
        print(f"测试文件不存在: {input_path}")
        return False

    with open(input_path, "rb") as f:
        file_content = base64.b64encode(f.read()).decode("utf-8")

    payload = json.dumps({
        "filename": os.path.basename(input_path),
        "file_content": file_content,
    }, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        f"{SERVICE_URL}/process",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    t0 = time.time()
    with urllib.request.urlopen(req, timeout=300) as resp:
        cost = time.time() - t0
        body = resp.read().decode("utf-8")
        result = json.loads(body)
        result["cost_seconds"] = round(cost, 3)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return result.get("success", False)


if __name__ == "__main__":
    if not test_health():
        print("健康检查失败")
        sys.exit(1)

    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # 默认找一个测试文件
        candidates = [
            r"D:\1\bom-server-test\B2604-4098马斌星_周周_.xls",
            r"D:\1\bom-server-test\RJCS_260418_006.xls",
        ]
        input_file = None
        for c in candidates:
            if os.path.exists(c):
                input_file = c
                break

    if not input_file or not os.path.exists(input_file):
        print("未找到测试文件，请提供 .xls 五金料单文件路径")
        sys.exit(1)

    print(f"测试文件: {input_file}")
    if test_process(input_file):
        print("\n[OK] /process 测试通过")
    else:
        print("\n[FAIL] /process 测试失败")
        sys.exit(1)
