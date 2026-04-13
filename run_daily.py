#!/usr/bin/env python3
"""
run_daily.py
────────────
Daily runner — refreshes Yahoo Finance valuations for all tickers in
the Tickers sheet and updates eToro_Master.xlsx Assumptions tab.

Steps:
  1. valuation.py  — fetches prices, runs DCF/DDM/EPV for every ticker
                     in the Tickers sheet (FTSE + international),
                     writes results to Assumptions and saves
                     reports/ftse_report.csv + reports/intl_report.csv
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
    result = subprocess.run([PYTHON, str(BASE_DIR / "scripts" / script)], env=os.environ)
    if result.returncode != 0:
        print(f"WARNING: {script} exited with code {result.returncode}")

if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Daily run starting...")
    load_env(ENV_FILE)
    run("valuation.py")
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Daily run complete.")
