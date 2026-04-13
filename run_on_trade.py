#!/usr/bin/env python3
"""
run_on_trade.py
───────────────
Run this script whenever you make a trade on eToro (buy or sell).
It compares your live eToro positions against eToro_Master.xlsx and
automatically applies any changes:

  New buy   → adds row to Portfolio tab, removes from Watchlist,
               updates Tickers flags (InPortfolio=Yes, InWatchlist=No)

  Full sell → removes row from Portfolio tab, adds to Closed Positions,
               moves back to Watchlist, updates Tickers flags

  Partial sell → reduces units/invested in Portfolio row,
                 splits dividends pro-rata and records sold portion
                 in Closed Positions tab

Notes:
  - Sale price is fetched from the eToro closed-positions API where
    available; otherwise the current market price is used as a fallback
    and the entry is flagged "(sale price estimated — please verify)".
  - Always check Closed Positions after running if you see the
    "(estimated)" flag, and update the sale price manually if needed.
  - Run run_daily.py afterwards to refresh valuations for any
    newly added tickers.
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
    result = subprocess.run(
        [PYTHON, str(BASE_DIR / "scripts" / script)],
        env=os.environ
    )
    if result.returncode != 0:
        print(f"WARNING: {script} exited with code {result.returncode}")
    return result.returncode


if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Trade sync starting...")
    load_env(ENV_FILE)

    rc = run("sync_portfolio.py")

    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Trade sync complete.")

    if rc == 0:
        print("\nTip: run run_daily.py next to refresh valuations for any new tickers.")
    else:
        print("\nWARNING: sync_portfolio.py reported errors — check output above.")
