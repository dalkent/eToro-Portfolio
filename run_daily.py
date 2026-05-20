#!/usr/bin/env python3
"""
run_daily.py
────────────
Daily / hourly runner. Refreshes valuations, regenerates the JSON cache, and
mirrors fresh files to Google Drive eToro Sync so every consumer (website,
tracker, dashboards, Mac) sees up-to-date data.

Steps:
  1. valuation.py            - fetches prices, runs DCF/DDM/EPV for every
                               ticker in the Tickers sheet, writes results to
                               eToro_Master.xlsx Assumptions, saves
                               reports/ftse_report.csv + reports/intl_report.csv.
  2. sync_xlsx_to_vault.py   - exports xlsx -> data/etoro_master.json (atomic
                               write, validated). The JSON is the canonical
                               source of truth for every downstream reader.
  3. mirror_to_drive.py      - copies xlsx + json + dashboard shells to
                               C:\\Users\\Neil\\My Drive\\eToro Sync\\ so the
                               Mac and other consumers see fresh data.
"""

import os
import subprocess
import sys
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
ENV_FILE = BASE_DIR / "etoro.env"
PYTHON   = sys.executable

def load_env(path: Path):
    """Load key=value pairs from a .env file into os.environ."""
    if not path.exists():
        print(f"WARNING: {path} not found")
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip())

def run(script: str):
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Running: {script}")
    # On Windows, suppress the child's cmd window when this script is invoked
    # from a hidden VBS / Task Scheduler context. CREATE_NO_WINDOW = 0x08000000.
    kwargs = {"env": os.environ}
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    result = subprocess.run([PYTHON, str(BASE_DIR / "scripts" / script)], **kwargs)
    if result.returncode != 0:
        print(f"WARNING: {script} exited with code {result.returncode}")
    return result.returncode

if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Daily run starting...")
    load_env(ENV_FILE)

    # 0) snapshot current xlsx + json BEFORE anything writes (safety net)
    run("backup_xlsx.py")

    # 1) valuations (writes the xlsx)
    run("valuation.py")

    # 2) regenerate JSON from xlsx (atomic write, validated)
    run("sync_xlsx_to_vault.py")

    # 3) mirror fresh files to Google Drive eToro Sync (xlsx, json, shells)
    run("mirror_to_drive.py")

    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Daily run complete.")
