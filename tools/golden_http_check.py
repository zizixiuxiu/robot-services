#!/usr/bin/env python3
"""Run golden-output checks against local HTTP file-processing services.

The script intentionally depends only on the Python standard library so it can
run on the server without preparing another environment.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
import sys
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError(f"{path} must contain a JSON object")
    return data


def _write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def _safe_name(value: str) -> str:
    value = value.strip() or "case"
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("._") or "case"


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _stable_digest(filename: str, data: bytes) -> tuple[str, str] | None:
    suffix = Path(filename).suffix.lower()
    if suffix not in {".zip", ".xlsx", ".xlsm"}:
        return None

    try:
        with zipfile.ZipFile(BytesIO(data), "r") as zf:
            digest = hashlib.sha256()
            for info in sorted(zf.infolist(), key=lambda item: item.filename):
                if info.is_dir():
                    continue
                entry_name = info.filename.replace("\\", "/")
                entry_data = zf.read(info)
                if entry_name == "docProps/core.xml":
                    entry_data = re.sub(
                        rb"<dcterms:(created|modified)[^>]*>.*?</dcterms:\1>",
                        rb"<dcterms:\1>__normalized__</dcterms:\1>",
                        entry_data,
                    )
                nested = _stable_digest(entry_name, entry_data)
                entry_hash = nested[1] if nested else _sha256_bytes(entry_data)
                digest.update(entry_name.encode("utf-8"))
                digest.update(b"\0")
                digest.update(entry_hash.encode("ascii"))
                digest.update(b"\0")
            return ("zip-content-sha256", digest.hexdigest())
    except zipfile.BadZipFile:
        return None


def _http_json(method: str, url: str, payload: dict[str, Any] | None = None, timeout: int = 300) -> tuple[int, Any]:
    body = None
    headers = {"Accept": "application/json"}
    if payload is not None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    req = Request(url, data=body, headers=headers, method=method)
    try:
        with urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            status = resp.status
    except HTTPError as exc:
        raw = exc.read()
        status = exc.code
    except URLError as exc:
        raise RuntimeError(f"request failed: {exc}") from exc

    try:
        parsed = json.loads(raw.decode("utf-8"))
    except Exception:
        parsed = {"raw": raw.decode("utf-8", errors="replace")}
    return status, parsed


def _file_payload(item: dict[str, Any]) -> dict[str, Any]:
    src = item.get("path") or item.get("input")
    if not src:
        raise ValueError("file item needs path or input")

    path = Path(src).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"input file not found: {path}")

    metadata_keys = {
        "fields",
        "health_path",
        "input",
        "name",
        "path",
        "process_path",
        "service",
        "service_url",
        "timeout_seconds",
        "url",
    }
    payload = {k: v for k, v in item.items() if k not in metadata_keys}
    payload.setdefault("filename", path.name)
    payload["file_content"] = base64.b64encode(path.read_bytes()).decode("ascii")
    return payload


def _build_payload(case: dict[str, Any]) -> dict[str, Any]:
    if "request" in case:
        request = case["request"]
        if not isinstance(request, dict):
            raise ValueError("case.request must be an object")
        return dict(request)

    fields = case.get("fields") or {}
    if not isinstance(fields, dict):
        raise ValueError("case.fields must be an object")
    payload = dict(fields)

    if "files" in case:
        files = case["files"]
        if not isinstance(files, list):
            raise ValueError("case.files must be a list")
        payload["files"] = [_file_payload(item) for item in files]
        return payload

    src = case.get("path") or case.get("input")
    if src:
        payload.update(_file_payload(case))
        return payload

    raise ValueError("case needs request, files, path, or input")


def _iter_output_items(response: Any) -> list[dict[str, Any]]:
    if not isinstance(response, dict):
        return []

    items: list[dict[str, Any]] = []

    if response.get("output_content"):
        items.append(
            {
                "filename": response.get("output_filename") or response.get("filename") or "output.bin",
                "file_content": response["output_content"],
                "source": "output_content",
            }
        )

    output_files = response.get("output_files")
    if isinstance(output_files, dict):
        for group_name, group_items in output_files.items():
            if isinstance(group_items, list):
                for item in group_items:
                    if isinstance(item, dict):
                        copied = dict(item)
                        copied.setdefault("group", group_name)
                        items.append(copied)
            elif isinstance(group_items, dict):
                copied = dict(group_items)
                copied.setdefault("group", group_name)
                items.append(copied)
    elif isinstance(output_files, list):
        for item in output_files:
            if isinstance(item, dict):
                items.append(dict(item))

    return items


def _decode_output_item(case_dir: Path, index: int, item: dict[str, Any]) -> dict[str, Any]:
    filename = str(item.get("filename") or f"output_{index}.bin")
    record: dict[str, Any] = {
        "index": index,
        "filename": filename,
        "group": item.get("group"),
        "path": item.get("path"),
    }

    content = item.get("file_content")
    if not content:
        record["missing_file_content"] = True
        record["error"] = item.get("error")
        return record

    try:
        data = base64.b64decode(content, validate=True)
    except Exception as exc:
        record["decode_error"] = str(exc)
        return record

    out_path = case_dir / f"{index:02d}_{_safe_name(filename)}"
    out_path.write_bytes(data)
    record.update(
        {
            "saved_as": str(out_path),
            "size": len(data),
            "sha256": _sha256_bytes(data),
        }
    )
    stable = _stable_digest(filename, data)
    if stable:
        record["stable_hash_type"] = stable[0]
        record["stable_sha256"] = stable[1]
    return record


def _run_case(case: dict[str, Any], output_root: Path) -> dict[str, Any]:
    name = str(case.get("name") or case.get("service") or "case")
    base_url = str(case.get("url") or case.get("service_url") or "").rstrip("/")
    if not base_url:
        raise ValueError(f"{name}: missing url")

    timeout = int(case.get("timeout_seconds") or 300)
    case_dir = output_root / _safe_name(name)
    case_dir.mkdir(parents=True, exist_ok=True)

    result: dict[str, Any] = {
        "name": name,
        "url": base_url,
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
    }

    health_path = str(case.get("health_path") or "/health")
    try:
        status, health = _http_json("GET", base_url + health_path, timeout=timeout)
        result["health"] = {"status_code": status, "body": health}
    except Exception as exc:
        result["health"] = {"error": str(exc)}

    payload = _build_payload(case)
    request_snapshot = _redact_payload(payload)
    _write_json(case_dir / "request.json", request_snapshot)

    process_path = str(case.get("process_path") or "/process")
    t0 = time.time()
    status, response = _http_json("POST", base_url + process_path, payload, timeout=timeout)
    elapsed = round(time.time() - t0, 3)

    result["process"] = {
        "status_code": status,
        "elapsed_seconds": elapsed,
        "success": response.get("success") if isinstance(response, dict) else None,
    }
    _write_json(case_dir / "response.json", response)

    outputs = []
    for idx, item in enumerate(_iter_output_items(response), start=1):
        outputs.append(_decode_output_item(case_dir, idx, item))
    result["outputs"] = outputs
    result["output_count"] = len(outputs)
    return result


def _redact_payload(value: Any) -> Any:
    if isinstance(value, dict):
        redacted = {}
        for key, item in value.items():
            if key == "file_content" and isinstance(item, str):
                redacted[key] = f"<base64:{len(item)} chars>"
            else:
                redacted[key] = _redact_payload(item)
        return redacted
    if isinstance(value, list):
        return [_redact_payload(item) for item in value]
    return value


def _compare(expected_path: Path, actual: dict[str, Any]) -> list[str]:
    expected = _read_json(expected_path)
    expected_cases = {case["name"]: case for case in expected.get("cases", [])}
    failures: list[str] = []

    for case in actual.get("cases", []):
        name = case.get("name")
        expected_case = expected_cases.get(name)
        if not expected_case:
            failures.append(f"{name}: missing in expected results")
            continue
        expected_outputs = expected_case.get("outputs", [])
        actual_outputs = case.get("outputs", [])
        expected_hashes = [_comparison_key(item) for item in expected_outputs]
        actual_hashes = [_comparison_key(item) for item in actual_outputs]
        if expected_hashes != actual_hashes:
            failures.append(f"{name}: output hashes changed")
    return failures


def _comparison_key(item: dict[str, Any]) -> tuple[Any, Any, Any, Any]:
    if item.get("stable_sha256") and item.get("stable_hash_type"):
        return (
            item.get("filename"),
            item.get("stable_hash_type"),
            item.get("stable_sha256"),
            None,
        )
    return (
        item.get("filename"),
        "raw-sha256",
        item.get("sha256"),
        item.get("size"),
    )


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, help="JSON manifest that describes service cases")
    parser.add_argument("--output-dir", default=None, help="Directory for decoded outputs and result summary")
    parser.add_argument("--compare", default=None, help="Optional previous results JSON to compare against")
    parser.add_argument("--case", action="append", default=None, help="Run only cases with this name. Can be repeated")
    args = parser.parse_args(argv)

    manifest_path = Path(args.manifest).expanduser()
    manifest = _read_json(manifest_path)
    cases = manifest.get("cases", [])
    if not isinstance(cases, list) or not cases:
        raise ValueError("manifest needs a non-empty cases list")
    if args.case:
        selected = set(args.case)
        cases = [case for case in cases if case.get("name") in selected]
        missing = selected - {case.get("name") for case in cases}
        if missing:
            raise ValueError(f"manifest does not contain selected case(s): {', '.join(sorted(missing))}")
        if not cases:
            raise ValueError("no cases selected")

    if args.output_dir:
        output_root = Path(args.output_dir).expanduser()
    else:
        output_root = Path("golden-results") / time.strftime("%Y%m%d-%H%M%S")
    output_root.mkdir(parents=True, exist_ok=True)

    summary: dict[str, Any] = {
        "manifest": str(manifest_path),
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "cases": [],
    }

    exit_code = 0
    for case in cases:
        try:
            summary["cases"].append(_run_case(case, output_root))
        except Exception as exc:
            exit_code = 1
            summary["cases"].append(
                {
                    "name": str(case.get("name") or "case"),
                    "error": str(exc),
                }
            )

    summary_path = output_root / "results.json"
    _write_json(summary_path, summary)

    if args.compare:
        failures = _compare(Path(args.compare).expanduser(), summary)
        summary["comparison"] = {"expected": args.compare, "failures": failures}
        _write_json(summary_path, summary)
        if failures:
            exit_code = 2
            for failure in failures:
                print(f"FAIL: {failure}", file=sys.stderr)

    ok_cases = sum(1 for case in summary["cases"] if not case.get("error"))
    print(f"Wrote {summary_path} ({ok_cases}/{len(cases)} cases completed)")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
