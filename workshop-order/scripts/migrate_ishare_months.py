# -*- coding: utf-8 -*-
"""一次性迁移：下车间归档加层级。四个方向：

    add-month    {姓名}/{原文件|处理后}/{文件}  →  {月}/{姓名}/{原文件|处理后}/{文件}
    undo-month   反向
    add-year     {月}/{姓名}/…                  →  {年}/{月}/{姓名}/…
    undo-year    反向

年份/月份取文件 mtime（三棵树的 mtime 都是北京时间，与处理时刻一致），不补零。

安全设计：
  * 默认 dry-run，必须显式 --apply 才真正搬。
  * 幂等：每个模式只认自己那一种深度的合规文件，已迁过的自动跳过，
    重复跑不会越搬越乱；中断后重跑即可收敛。
  * 只搬不删：同盘 shutil.move（本地是 rename，UNC 是 SMB 服务端 rename），
    遍历用 os.walk 且剪掉过深子树。
  * 空目录自底向上 rmdir —— rmdir 只对空目录成功，非空一律留着。
  * test/ 整棵跳过；不合规的零散文件（如 Thumbs.db）原地不动。
  * 目录冲突时保留双方并打 CONFLICT（不静默覆盖）；与目标 size+mtime 一致的
    残留判为等价副本，删源收尾。

用法（Git Bash 下必须 MSYS_NO_PATHCONV=1，中文路径才不会被转换）：
    python migrate_ishare_months.py --mode add-year                # dry-run
    python migrate_ishare_months.py --mode add-year --apply
    python migrate_ishare_months.py --manifest before.json         # 出归一化清单
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

STAGE_ROOT = r"D:\Services\robot-services\workshop-order\data\ishare-archive"
BACKUP_ROOT = r"D:\Ishare_2026_backup"
UNC_ROOT = r"\\192.168.8.135\Ishare录单临时文件\2026"

# 迁移顺序：先本地（快、能最快暴露脚本问题），再远程。
ROOTS = [
    ("BACKUP", BACKUP_ROOT),
    ("STAGE", STAGE_ROOT),
    ("UNC", UNC_ROOT),
]
SKIP_TOP = {"test"}
TYPE_DIRS = {"原文件", "处理后"}


def _stdout_utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
        except (AttributeError, OSError):
            pass


def _is_month(seg: str) -> bool:
    return seg.isdigit() and 1 <= int(seg) <= 12


def _is_year(seg: str) -> bool:
    return seg.isdigit() and 1000 <= int(seg) <= 2999


def _year_of(st: os.stat_result) -> str:
    return str(datetime.fromtimestamp(st.st_mtime).year)


def _month_of(st: os.stat_result) -> str:
    return str(datetime.fromtimestamp(st.st_mtime).month)


# ---- 各模式的目标路径推导：返回 None 表示这个文件不归本模式管 ----

def _add_month(rel: Path, st: os.stat_result) -> Path | None:
    if rel.parts[1] not in TYPE_DIRS or _is_month(rel.parts[0]) or _is_year(rel.parts[0]):
        return None
    return Path(_month_of(st)) / rel


def _undo_month(rel: Path, st: os.stat_result) -> Path | None:
    if not _is_month(rel.parts[0]) or rel.parts[2] not in TYPE_DIRS:
        return None
    return Path(*rel.parts[1:])


def _add_year(rel: Path, st: os.stat_result) -> Path | None:
    if not _is_month(rel.parts[0]) or rel.parts[2] not in TYPE_DIRS:
        return None
    return Path(_year_of(st)) / rel


def _undo_year(rel: Path, st: os.stat_result) -> Path | None:
    if not _is_year(rel.parts[0]) or not _is_month(rel.parts[1]):
        return None
    if rel.parts[3] not in TYPE_DIRS:
        return None
    return Path(*rel.parts[1:])


# 模式 → (要看的文件深度, 目标推导, 迁完后哪种目录算作废可清)
MODES = {
    "add-month":  (3, _add_month,  "flat"),
    "undo-month": (4, _undo_month, "month"),
    "add-year":   (4, _add_year,   "month"),
    "undo-year":  (5, _undo_year,  "year"),
}


def _iter_files(root: Path, depth: int):
    """产出 (rel, src)：rel 段数恰好等于 depth；test/ 剪枝，过深子树不进去。"""
    for dirpath, dirnames, filenames in os.walk(root):
        here = Path(dirpath)
        rel_dir = here.relative_to(root)
        if rel_dir.parts and rel_dir.parts[0] in SKIP_TOP:
            dirnames[:] = []
            continue
        if len(rel_dir.parts) > depth:
            dirnames[:] = []
            continue
        for name in filenames:
            rel = rel_dir / name
            if len(rel.parts) == depth:
                yield rel, here / name


def relocate(root: Path, applied: bool, depth: int, target_of) -> tuple[int, int, list[str], set[Path]]:
    moved = skipped = 0
    notes: list[str] = []
    sources: set[Path] = set()
    for rel, src in _iter_files(root, depth):
        st = src.stat()
        target_rel = target_of(rel, st)
        if target_rel is None:
            continue
        dst = root / target_rel
        if dst.exists():
            d = dst.stat()
            if st.st_size == d.st_size and abs(st.st_mtime - d.st_mtime) < 1:
                # 上次中断留下的等价副本 → 收尾删源
                if applied:
                    src.unlink()
                sources.add(src)
                notes.append(f"DUP-EQ   {rel}  （与目标等价，删源）")
                skipped += 1
            else:
                notes.append(f"CONFLICT {rel}  ←→  {target_rel}  （保留双方，待人工看）")
                skipped += 1
            continue
        notes.append(f"{rel}  ->  {target_rel}")
        sources.add(src)
        if applied:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
        moved += 1
    return moved, skipped, notes, sources


def prune_empty_dirs(root: Path, applied: bool, moved_sources: set[Path], stale: str) -> list[Path]:
    """删掉搬空了的作废目录。

    stale 指明「哪种顶层段算过时」：flat=旧扁平位（既不是年也不是月），
    month=月份层的残留，year=年份层的残留。反向的那一层（比如正向迁移后
    新出现的月份夹）一律不碰 —— 容器随时可能正在往里写，删掉瞬时为空的
    目录会和它抢跑。rmdir 对非空目录必然失败，不存在误删。
    dry-run 时用 moved_sources 推演「搬完之后谁会变空」，让预览数字与实际一致。
    """
    keep: set[Path] = set()
    for dirpath, _dirnames, filenames in os.walk(root):
        here = Path(dirpath)
        for name in filenames:
            if (here / name) in moved_sources:
                continue          # 这个文件将被搬走，不算数
            keep.add(here)
            for anc in here.parents:
                if anc == root:
                    break
                keep.add(anc)

    def is_stale(seg: str) -> bool:
        if stale == "flat":
            return not _is_month(seg) and not _is_year(seg)
        if stale == "month":
            return _is_month(seg)
        return _is_year(seg)

    removed: list[Path] = []
    for dirpath, _dirnames, _filenames in os.walk(root, topdown=False):
        here = Path(dirpath)
        if here == root or here in keep:
            continue
        parts = here.relative_to(root).parts
        if SKIP_TOP & set(parts) or not is_stale(parts[0]):
            continue
        if applied:
            try:
                here.rmdir()
            except OSError:
                continue
        removed.append(here)
    return removed


def manifest(root: Path) -> list[dict]:
    """按 (姓名, 类型, 文件名) 归一化的清单 —— 与目录层级无关，
    扁平 / 带月份 / 带年份的布局都能比对。迁移前后各出一份用于核对。"""
    rows = []
    for dirpath, dirnames, filenames in os.walk(root):
        here = Path(dirpath)
        rel_dir = here.relative_to(root)
        if rel_dir.parts and rel_dir.parts[0] in SKIP_TOP:
            dirnames[:] = []
            continue
        parts = rel_dir.parts
        if len(parts) < 2 or parts[-1] not in TYPE_DIRS:
            continue
        for name in filenames:
            st = (here / name).stat()
            rows.append({
                "person": parts[-2],
                "type": parts[-1],
                "name": name,
                "size": st.st_size,
                "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            })
    return sorted(rows, key=lambda r: (r["person"], r["type"], r["name"]))


def main() -> int:
    _stdout_utf8()
    ap = argparse.ArgumentParser(description="下车间归档层级迁移（默认 dry-run）")
    ap.add_argument("--mode", choices=sorted(MODES), default="add-month")
    ap.add_argument("--apply", action="store_true", help="真正执行；不加则只预览")
    ap.add_argument("--only", action="append", choices=[n for n, _ in ROOTS],
                    help="只处理指定树，可重复；默认全部")
    ap.add_argument("--root", action="append", metavar="NAME=PATH",
                    help="覆盖某棵树的位置（如 --root UNC=\\\\host\\share），可重复")
    ap.add_argument("--manifest", metavar="PATH", help="把归一化文件清单写成 JSON")
    args = ap.parse_args()

    roots = list(ROOTS)
    for item in args.root or []:
        name, sep, path = item.partition("=")
        if not sep or name not in dict(ROOTS):
            ap.error(f"--root 需要 NAME=PATH 且 NAME ∈ {[n for n, _ in ROOTS]}，收到 {item!r}")
        roots = [(n, path if n == name else p) for n, p in roots]

    targets = [(n, p) for n, p in roots if not args.only or n in args.only]
    depth, target_of, stale = MODES[args.mode]
    print(f"=== {args.mode} {'APPLY' if args.apply else 'DRY-RUN'} ===")

    for label, raw in targets:
        root = Path(raw)
        print(f"\n--- {label}: {root}")
        if not root.exists():
            print("    根目录不存在，跳过")
            continue
        moved, skipped, notes, sources = relocate(root, args.apply, depth, target_of)
        for line in notes[:40]:
            print(f"    {line}")
        if len(notes) > 40:
            print(f"    ...（另有 {len(notes) - 40} 条）")
        removed = prune_empty_dirs(root, args.apply, sources, stale)
        print(f"    {'已移动' if args.apply else '将移动'} {moved} 个；"
              f"跳过 {skipped} 条；{'已删' if args.apply else '可删'}空目录 {len(removed)} 个")
        if args.manifest:
            out = Path(args.manifest)
            payload = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
            payload[label] = manifest(root)
            out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"    清单已写入 {out}（{label} 共 {len(payload[label])} 条）")

    if not args.apply:
        print("\n（dry-run，未改动任何文件；确认无误后加 --apply）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
