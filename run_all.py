#!/usr/bin/env python3
"""
run_all.py
──────────
Single command to refresh everything and regenerate the dashboard.

Steps:
  1. sync_portfolio.py  — pulls live positions from eToro API, syncs eToro_Master.xlsx
  2. valuation.py       — fetches prices, runs DCF/DDM/EPV, updates Assumptions sheet
  3. generate_dashboard.py — reads Excel + live prices, writes eToro_dashboard.html

Usage:
  python run_all.py            # full run (sync + valuations + dashboard)
  python run_all.py --dash     # skip sync & valuations, just regenerate dashboard
  python run_all.py --no-sync  # skip eToro API sync, run valuations + dashboard only

Output:
  eToro_dashboard.html — open in any browser. No server required.
"""

import argparse
import os
import subprocess
import sys
import webbrowser
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
ENV_FILE = BASE_DIR / "etoro.env"
SCRIPTS  = BASE_DIR / "scripts"
PYTHON   = sys.executable

sys.path.insert(0, str(SCRIPTS))
from paths import UPLOAD_DIR, SYNC_DIR


def load_env(path: Path):
    if not path.exists():
        print(f"  WARNING: {path} not found — eToro API calls may fail")
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip())


def run(label: str, script: Path) -> bool:
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] {label}")
    result = subprocess.run([PYTHON, str(script)], env=os.environ)
    if result.returncode != 0:
        print(f"  WARNING: exited with code {result.returncode}")
        return False
    return True


def main():
    parser = argparse.ArgumentParser(description="Refresh eToro dashboard")
    parser.add_argument("--dash",    action="store_true", help="Regenerate dashboard only (no API calls)")
    parser.add_argument("--no-sync", action="store_true", help="Skip eToro portfolio sync, run valuations + dashboard")
    parser.add_argument("--open",    action="store_true", help="Open eToro_dashboard.html in browser when done")
    args = parser.parse_args()

    print("=" * 56)
    print("  run_all.py — eToro Dashboard Refresh")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 56)

    load_env(ENV_FILE)

    if not args.dash:
        if not args.no_sync:
            run("Step 1/3 — Syncing portfolio from eToro API ...", SCRIPTS / "sync_portfolio.py")
        else:
            print("\n  Skipping Step 1 (--no-sync)")

        run("Step 2/3 — Running valuations (DCF/DDM/EPV) ...", SCRIPTS / "valuation.py")

    run("Step 3/4 — Generating dashboard ...", SCRIPTS / "generate_dashboard.py")

    run("Step 4/4 — Exporting xlsx to vault + JSON cache ...", SCRIPTS / "sync_xlsx_to_vault.py")

    dashboard = BASE_DIR / "dashboards" / "eToro_dashboard.html"
    if dashboard.exists():  # copy regardless of exit code — partial dashboard is better than none
        print(f"\n  Dashboard ready -> {dashboard}")
        # Copy to shared Upload folder regardless of whether price fetch succeeded
        import shutil
        upload_dest = UPLOAD_DIR / "eToro_dashboard.html"
        try:
            upload_dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(dashboard), str(upload_dest))
            print(f"  Copied to → {upload_dest}")
        except Exception as e:
            print(f"  WARNING: could not copy to Upload folder — {e}")
        if args.open:
            webbrowser.open(dashboard.as_uri())
            print("  Opened in browser.")
    else:
        print("\n  WARNING: eToro_dashboard.html not found — check for errors above")

    # ── Mirror the freshly-revalued workbook to Drive so the Mac can pick it up.
    # PC is the data master; Mac is read-only on the xlsx.
    import shutil as _sh
    src_xlsx = BASE_DIR / "data" / "eToro_Master.xlsx"
    drive_xlsx = SYNC_DIR / "eToro_Master.xlsx"
    if src_xlsx.exists():
        try:
            drive_xlsx.parent.mkdir(parents=True, exist_ok=True)
            _sh.copy2(str(src_xlsx), str(drive_xlsx))
            print(f"  Mirrored xlsx to Drive -> {drive_xlsx}")
        except Exception as e:
            print(f"  WARNING: could not mirror xlsx to Drive — {e}")

    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Done.\n")


if __name__ == "__main__":
    main()
