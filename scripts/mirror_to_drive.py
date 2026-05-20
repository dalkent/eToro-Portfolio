#!/usr/bin/env python3
"""
mirror_to_drive.py
──────────────────
Mirror the canonical data files from C:\\Users\\Neil\\ClaudeCode\\eToro to the
Google Drive 'eToro Sync' folder so the Mac (and any other reader) sees fresh
data without needing a GitHub pull.

Mirrored files:
  data/eToro_Master.xlsx        - upstream master spreadsheet
  data/etoro_master.json        - canonical JSON cache (source of truth for readers)
  data/bookmarks.json           - manually curated, PC only
  dashboards/dashboard2.html    - briefing source shell (hand-crafted)
  dashboards/health_dashboard.html - health source shell (hand-crafted)

This script is idempotent and safe to call from any BAT or python runner.
It exits 0 if all files mirror successfully, or if a source file simply
doesn't exist yet (it warns and skips). It exits 1 only on actual copy
failures (permission, Drive offline, etc.).

Designed to be called at the END of any pipeline that updates the xlsx or
JSON, so consumers always see fresh data.

Usage:
    python scripts/mirror_to_drive.py
    python scripts/mirror_to_drive.py --quiet
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

# paths.py lives in the same scripts/ dir; resolve drive paths via it so this
# module works on Windows and macOS without hardcoded paths.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import REPO_DIR, SYNC_DIR  # noqa: E402

BASE_DIR  = REPO_DIR
DRIVE_DIR = SYNC_DIR

# (source subdir relative to BASE_DIR, filename)
MIRROR_FILES = [
    ("data",       "eToro_Master.xlsx"),
    ("data",       "etoro_master.json"),
    ("data",       "bookmarks.json"),
    ("dashboards", "dashboard2.html"),
    ("dashboards", "health_dashboard.html"),
]


def mirror(quiet: bool = False) -> int:
    """Copy each mirror file from local to Drive. Returns count of successful copies."""
    ok = 0
    fail = 0
    skip = 0
    try:
        DRIVE_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"ERROR: cannot create Drive folder {DRIVE_DIR}: {e}", file=sys.stderr)
        return -1

    for src_subdir, fname in MIRROR_FILES:
        src = BASE_DIR / src_subdir / fname
        dest = DRIVE_DIR / fname
        if not src.exists():
            if not quiet:
                print(f"  SKIP {fname}: source {src} not found")
            skip += 1
            continue
        try:
            shutil.copy2(str(src), str(dest))
            ok += 1
            if not quiet:
                print(f"  OK   {fname}: {src_subdir}/{fname} -> Drive")
        except Exception as e:
            print(f"  FAIL {fname}: {e}", file=sys.stderr)
            fail += 1

    if not quiet:
        print(f"\nmirror_to_drive: ok={ok} skip={skip} fail={fail}")

    return -1 if fail else ok


def main() -> int:
    parser = argparse.ArgumentParser(description="Mirror eToro data files to Google Drive eToro Sync.")
    parser.add_argument("--quiet", action="store_true", help="Only print on failure.")
    args = parser.parse_args()
    rc = mirror(quiet=args.quiet)
    return 1 if rc < 0 else 0


if __name__ == "__main__":
    sys.exit(main())
