#!/usr/bin/env python3
"""
run_combined.py
───────────────
Refreshes the combined eToro + Trading 212 ISA dashboard.

Steps:
  1. sync_t212.py                 — pulls T212 portfolio + cash
  2. generate_combined_dashboard.py — writes t212_dashboard.html

Does NOT touch the existing eToro pipeline. The eToro section reads whatever
is already in data/eToro_Master.xlsx — so run_all.py should be run first (or
on its normal schedule) to keep eToro data fresh.

Usage:
    python run_combined.py          # sync T212 + regenerate combined dashboard
    python run_combined.py --dash   # skip T212 sync, just regenerate HTML
    python run_combined.py --open   # open t212_dashboard.html when done
"""

import argparse
import os
import shutil
import subprocess
import sys
import webbrowser
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
SCRIPTS  = BASE_DIR / "scripts"
PYTHON   = sys.executable

ENV_FILES = [BASE_DIR / "etoro.env", BASE_DIR / "t212.env"]
UPLOAD_DEST = Path(r"C:\Users\Neil\My Drive\Upload\t212_dashboard.html")


def load_env(path: Path):
    if not path.exists():
        print(f"  note: {path.name} not found — skipping")
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip())


def run(label: str, script: Path) -> int:
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] {label}")
    return subprocess.run([PYTHON, str(script)], env=os.environ).returncode


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dash", action="store_true", help="Skip T212 sync")
    ap.add_argument("--open", action="store_true", help="Open HTML in browser")
    args = ap.parse_args()

    print("=" * 56)
    print("  run_combined.py — eToro + T212 ISA dashboard")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 56)

    for e in ENV_FILES:
        load_env(e)

    if not args.dash:
        run("Step 1/2 — Syncing Trading 212 ...", SCRIPTS / "sync_t212.py")
    else:
        print("\n  Skipping T212 sync (--dash)")

    run("Step 2/2 — Generating combined dashboard ...",
        SCRIPTS / "generate_combined_dashboard.py")

    html = BASE_DIR / "dashboards" / "t212_dashboard.html"
    if html.exists():
        print(f"\n  Dashboard ready -> {html}")
        try:
            UPLOAD_DEST.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(html), str(UPLOAD_DEST))
            print(f"  Copied to    -> {UPLOAD_DEST}")
        except Exception as e:
            print(f"  WARNING: could not copy to Upload folder — {e}")
        if args.open:
            webbrowser.open(html.as_uri())
    else:
        print("\n  WARNING: t212_dashboard.html not produced — see errors above")


if __name__ == "__main__":
    main()
