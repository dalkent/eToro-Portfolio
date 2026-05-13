#!/usr/bin/env python3
"""
run_macro.py
────────────
Refreshes the macro dashboard.
  1. sync_macro.py                    — pulls FRED + yfinance data
  2. generate_macro_dashboard.py      — writes macro_dashboard.html
  3. Copies to C:\\Users\\Neil\\My Drive\\Upload\\macro_dashboard.html
  4. Copies to ..\\Homelab\\static-html\\macro_dashboard.html so Homepage can serve it

Usage:
    python run_macro.py          # full refresh
    python run_macro.py --dash   # skip sync, just regenerate HTML
    python run_macro.py --open   # open macro_dashboard.html when done
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

sys.path.insert(0, str(SCRIPTS))
from paths import UPLOAD_DIR

ENV_FILES = [BASE_DIR / "etoro.env", BASE_DIR / "t212.env"]

UPLOAD_DEST   = UPLOAD_DIR / "macro_dashboard.html"
HOMELAB_DEST  = BASE_DIR.parent / "Homelab" / "static-html" / "macro_dashboard.html"


def load_env(path: Path):
    if not path.exists():
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            os.environ.setdefault(k.strip(), v.strip())


def run(label: str, script: Path) -> int:
    print(f"\n[{datetime.now():%H:%M:%S}] {label}")
    return subprocess.run([PYTHON, str(script)], env=os.environ).returncode


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dash", action="store_true", help="Skip data sync")
    ap.add_argument("--open", action="store_true", help="Open HTML when done")
    args = ap.parse_args()

    print("=" * 56)
    print("  run_macro.py — macro + markets dashboard")
    print(f"  {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 56)

    for e in ENV_FILES:
        load_env(e)

    if not args.dash:
        run("Step 1/2 - Syncing macro data ...", SCRIPTS / "sync_macro.py")
    else:
        print("\n  Skipping macro sync (--dash)")

    run("Step 2/2 - Generating macro dashboard ...", SCRIPTS / "generate_macro_dashboard.py")

    html = BASE_DIR / "dashboards" / "macro_dashboard.html"
    if html.exists():
        print(f"\n  Dashboard ready -> {html}")
        for dest in (UPLOAD_DEST, HOMELAB_DEST):
            try:
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(str(html), str(dest))
                print(f"  Copied to    -> {dest}")
            except Exception as e:
                print(f"  WARNING: could not copy to {dest.parent.name} - {e}")
        if args.open:
            webbrowser.open(html.as_uri())
    else:
        print("\n  WARNING: macro_dashboard.html not produced")


if __name__ == "__main__":
    main()
