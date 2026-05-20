#!/usr/bin/env python3
"""
backup_xlsx.py
──────────────
Take a timestamped backup of eToro_Master.xlsx into data/backups/ and prune
anything older than RETENTION_DAYS. Designed to be called by run_daily.py /
eToro.bat / website.bat --revalue BEFORE valuation.py runs, so if a write
corrupts the file we always have a recent snapshot.

The backup is a straight byte-for-byte copy (shutil.copy2) so it preserves
the original mtime. It's saved under data/backups/ inside the eToro repo
(gitignored by the existing data/ rule).

Usage:
    python scripts/backup_xlsx.py
    python scripts/backup_xlsx.py --retention 30   (override retention days)
    python scripts/backup_xlsx.py --quiet
"""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR    = Path(__file__).parent.parent
DATA_DIR    = BASE_DIR / "data"
BACKUP_DIR  = DATA_DIR / "backups"
LIVE_XLSX   = DATA_DIR / "eToro_Master.xlsx"
LIVE_JSON   = DATA_DIR / "etoro_master.json"

RETENTION_DAYS = 14
BACKUP_PREFIX  = "eToro_Master."


def take_backup(quiet: bool = False) -> Path | None:
    """Snapshot the current xlsx into data/backups/. Returns the new backup path
    or None if there was nothing to back up."""
    if not LIVE_XLSX.exists():
        if not quiet:
            print(f"  SKIP backup: {LIVE_XLSX} not found")
        return None
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    dest = BACKUP_DIR / f"{BACKUP_PREFIX}{stamp}.xlsx"
    shutil.copy2(str(LIVE_XLSX), str(dest))
    if not quiet:
        size_kb = dest.stat().st_size / 1024
        print(f"  Backed up xlsx -> {dest.name} ({size_kb:.0f} KB)")
    # Also snapshot the JSON if it exists - cheap insurance.
    if LIVE_JSON.exists():
        json_dest = BACKUP_DIR / f"etoro_master.{stamp}.json"
        shutil.copy2(str(LIVE_JSON), str(json_dest))
        if not quiet:
            size_kb = json_dest.stat().st_size / 1024
            print(f"  Backed up json -> {json_dest.name} ({size_kb:.0f} KB)")
    return dest


def prune_old(retention_days: int, quiet: bool = False) -> int:
    """Delete backups older than retention_days. Returns count deleted."""
    if not BACKUP_DIR.exists():
        return 0
    cutoff = datetime.now() - timedelta(days=retention_days)
    deleted = 0
    for p in BACKUP_DIR.iterdir():
        if not p.is_file():
            continue
        if not (p.name.startswith(BACKUP_PREFIX) or p.name.startswith("etoro_master.")):
            continue
        try:
            mtime = datetime.fromtimestamp(p.stat().st_mtime)
        except OSError:
            continue
        if mtime < cutoff:
            try:
                p.unlink()
                deleted += 1
                if not quiet:
                    print(f"  Pruned old backup: {p.name}")
            except OSError as e:
                print(f"  WARN: could not prune {p.name}: {e}", file=sys.stderr)
    return deleted


def main() -> int:
    parser = argparse.ArgumentParser(description="Backup eToro_Master.xlsx + etoro_master.json with rolling retention.")
    parser.add_argument("--retention", type=int, default=RETENTION_DAYS,
                        help=f"Days to keep backups (default: {RETENTION_DAYS})")
    parser.add_argument("--quiet", action="store_true", help="Only print on action.")
    args = parser.parse_args()

    backup = take_backup(quiet=args.quiet)
    deleted = prune_old(args.retention, quiet=args.quiet)
    if not args.quiet:
        print(f"backup_xlsx: backup={'yes' if backup else 'no'} pruned={deleted}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
