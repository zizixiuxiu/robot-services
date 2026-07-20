import argparse
import os
import json
import subprocess
import sys
import time
import socket
from datetime import datetime
from pathlib import Path

import pyodbc

from sqlserver_readonly import _build_connection_string, _load_config


BASE_DIR = Path(os.getenv("ORDER_FLOW_BASE_DIR", r"D:\Services\robot-services\order-flow-monitor"))
WORKDIR = Path(os.getenv("ORDER_FLOW_WORKDIR", str(BASE_DIR / "src")))
PYTHON_EXE = os.getenv(
    "ORDER_FLOW_PYTHON_EXE",
    r"C:\Users\Administrator\.workbuddy\binaries\python\envs\bom-server\Scripts\python.exe",
)
STATE_PATH = Path(os.getenv("ORDER_FLOW_STATE_PATH", str(BASE_DIR / "data" / "order_flow_refresh_state.json")))
LOG_PATH = Path(os.getenv("ORDER_FLOW_LOG_PATH", str(BASE_DIR / "logs" / "order_flow_refresh.log")))
SINGLE_INSTANCE_PORT = int(os.getenv("ORDER_FLOW_SINGLE_INSTANCE_PORT", "61235"))
_single_instance_socket = None


CHECK_SQL = """
WITH BaseOrders AS (
    SELECT o.Id, o.OrderId, o.Lchange, o.SalesOrder_Id
    FROM dbo.T_BOM_Order AS o
    WHERE o.Lchange >= '2026-05-01'
      AND o.Lchange < DATEADD(day, 1, CONVERT(date, GETDATE()))
),
OrderSig AS (
    SELECT
        COUNT_BIG(*) AS OrderCnt,
        MAX(Lchange) AS MaxLchange,
        CHECKSUM_AGG(CHECKSUM(OrderId, Lchange, SalesOrder_Id)) AS OrderChecksum
    FROM BaseOrders
),
DetailSig AS (
    SELECT
        COUNT_BIG(*) AS DetailCnt,
        CHECKSUM_AGG(CHECKSUM(
            d.Id,
            d.OrderId,
            d.Product_Id,
            d.SysType,
            d.Fthk,
            d.MatProducer,
            d.Matname,
            d.Surftnam,
            d.DetailType,
            d.DetailName,
            d.Info4,
            d.AssemblyArea,
            d.Cnt
        )) AS DetailChecksum
    FROM BaseOrders AS o
    JOIN dbo.T_BOM_Item AS b ON b.Order_Id = o.Id
    JOIN dbo.T_BOM_ItemDetail AS d ON d.Product_Id = b.Id
)
SELECT
    o.OrderCnt,
    o.MaxLchange,
    o.OrderChecksum,
    d.DetailCnt,
    d.DetailChecksum
FROM OrderSig AS o
CROSS JOIN DetailSig AS d
"""


def log(message):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message}"
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        # Windows console may use a limited codepage (e.g. GBK).
        # Write raw UTF-8 bytes to stdout.buffer to bypass the text encoding layer.
        sys.stdout.buffer.write(line.encode("utf-8", errors="replace") + b"\n")
        sys.stdout.buffer.flush()
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def ensure_single_instance():
    global _single_instance_socket
    _single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        _single_instance_socket.bind(("127.0.0.1", SINGLE_INSTANCE_PORT))
        _single_instance_socket.listen(1)
    except OSError as exc:
        log(f"Another OrderFlowMonitor instance is already running on port {SINGLE_INSTANCE_PORT}: {exc}")
        sys.exit(0)


def load_state():
    if not STATE_PATH.exists():
        return None
    return json.loads(STATE_PATH.read_text(encoding="utf-8"))


def save_state(signature):
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(signature, ensure_ascii=False, indent=2), encoding="utf-8")


def current_signature():
    conn = pyodbc.connect(_build_connection_string(_load_config()), timeout=30)
    try:
        cur = conn.cursor()
        cur.execute(CHECK_SQL)
        row = cur.fetchone()
        return {
            "order_count": int(row.OrderCnt or 0),
            "max_lchange": row.MaxLchange.isoformat(sep=" ") if row.MaxLchange else None,
            "order_checksum": int(row.OrderChecksum) if row.OrderChecksum is not None else None,
            "detail_count": int(row.DetailCnt or 0),
            "detail_checksum": int(row.DetailChecksum) if row.DetailChecksum is not None else None,
        }
    finally:
        conn.close()


def classify_refresh_failure(output):
    if "reason=wps_workbook_not_stable" in output or "did not become stable" in output:
        return "wps_not_stable_skipped"
    if "reason=workbook_changed_during_refresh" in output or "Target workbook changed while refresh was running" in output:
        return "wps_changed_during_refresh_skipped"
    if "reason=target_workbook_not_found" in output or "Target workbook not found" in output:
        return "target_workbook_not_found"
    if "No target workbook was refreshed" in output:
        return "no_target_refreshed"
    return "refresh_failed"


def run_logged(command, label):
    result = subprocess.run(
        command,
        cwd=WORKDIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    output = result.stdout.decode("utf-8", errors="replace")
    for line in output.splitlines():
        log(f"{label}: {line}")
    if result.returncode != 0:
        reason = classify_refresh_failure(output)
        raise RuntimeError(f"{label} failed reason={reason} exit_code={result.returncode}")
    return output


def refresh_workbook():
    log("Refresh step=fetch_sql status=started")
    run_logged(
        [PYTHON_EXE, "fetch_sql_data_wsl.py"],
        "fetch_sql",
    )

    log("Refresh step=append_workbook status=started")
    run_logged(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(WORKDIR / "append_new_orders_to_formal.ps1"),
        ],
        "append_workbook",
    )
    log("Refresh status=complete reason=database_changed")


def run_once(force=False):
    signature = current_signature()
    previous = load_state()
    if force or previous != signature:
        change_reason = "forced" if force else "database_changed"
        log(f"Refresh needed reason={change_reason} previous={previous} current={signature}")
        refresh_workbook()
        save_state(signature)
        return True
    log(f"Refresh skipped reason=database_no_change current={signature}")
    return False


def main():
    ensure_single_instance()

    parser = argparse.ArgumentParser(description="Monitor SQL Server order-flow changes and refresh workbook when needed.")
    parser.add_argument("--once", action="store_true", help="Run one check and exit.")
    parser.add_argument("--force", action="store_true", help="Force a refresh regardless of state.")
    parser.add_argument("--interval", type=int, default=180, help="Polling interval in seconds. Default: 180.")
    args = parser.parse_args()

    if args.once:
        run_once(force=args.force)
        return

    log(f"Monitor started. interval={args.interval}s")
    while True:
        try:
            run_once(force=args.force)
            args.force = False
        except Exception as exc:
            log(f"ERROR: {exc}")
        time.sleep(args.interval)


if __name__ == "__main__":
    main()
