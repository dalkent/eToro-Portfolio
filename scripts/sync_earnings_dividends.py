"""Fetch per-ticker earnings + dividend info via Yahoo, write JSON for the dashboard.

Replaces the inline per-refresh Yahoo sweeps (which cost ~4.7s per dashboard
refresh). Scheduled task runs this once a day — data only changes slowly.

Writes:
  data/earnings.json    — same shape as the old inline earnings_calendar() output
  data/dividends.json   — same shape as the old inline dividend_calendar() output
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent.parent
sys.path.insert(0, str(BASE))

import run_news_server as srv  # noqa: E402

DATA_DIR = BASE / "data"


def main() -> None:
    srv._cache.clear()
    earnings = srv.earnings_calendar()
    dividends = srv.dividend_calendar()

    now_iso = datetime.now().astimezone().isoformat(timespec="seconds")

    earnings_path = DATA_DIR / "earnings.json"
    earnings_path.write_text(json.dumps({
        "generated_at": now_iso,
        "items":        earnings,
    }, indent=2), encoding="utf-8")

    dividends_path = DATA_DIR / "dividends.json"
    dividends_path.write_text(json.dumps({
        "generated_at": now_iso,
        "items":        dividends,
    }, indent=2), encoding="utf-8")

    print(f"Wrote {len(earnings)} earnings -> {earnings_path}")
    print(f"Wrote {len(dividends)} dividends -> {dividends_path}")


if __name__ == "__main__":
    main()
