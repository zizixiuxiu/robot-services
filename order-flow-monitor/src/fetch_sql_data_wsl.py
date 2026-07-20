import json
from decimal import Decimal
import glob
import importlib.util
from pathlib import Path
import time

import pyodbc


WORKDIR = Path(r"C:\Users\Administrator\Documents\Codex\2026-05-28\sqlserver")
SQL_PATH = WORKDIR / "optimized_sqlserver_query.sql"
OUTPUT_JSON = WORKDIR / "sql_data.json"


def load_sqlserver_config_module():
    # Look in the same directory as this script first, then fall back to common locations.
    candidates = [
        Path(__file__).resolve().parent / "sqlserver_readonly.py",
        WORKDIR / "sqlserver_readonly.py",
        Path(r"D:\1\sqlserver-monitor\sqlserver_readonly.py"),
    ]
    for candidate in candidates:
        if candidate.exists():
            spec = importlib.util.spec_from_file_location("sqlserver_readonly", str(candidate))
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module
    raise FileNotFoundError(
        f"sqlserver_readonly.py not found in any of: {[str(c) for c in candidates]}"
    )


def convert(value):
    if value is None:
        return None
    if isinstance(value, Decimal):
        return float(value)
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


sql_config = load_sqlserver_config_module()
sql = SQL_PATH.read_text(encoding="utf-8-sig")
conn = pyodbc.connect(sql_config._build_connection_string(sql_config._load_config()), timeout=180)
try:
    cur = conn.cursor()
    start = time.time()
    cur.execute(sql)
    headers = [d[0] for d in cur.description]
    rows = [[convert(v) for v in row] for row in cur.fetchall()]
    elapsed_ms = int((time.time() - start) * 1000)
finally:
    conn.close()

OUTPUT_JSON.write_text(
    json.dumps(
        {"headers": headers, "rows": rows, "elapsed_ms": elapsed_ms},
        ensure_ascii=False,
    ),
    encoding="utf-8",
)
print(f"output={OUTPUT_JSON}")
print(f"rows={len(rows)}")
print(f"sql_elapsed_ms={elapsed_ms}")
