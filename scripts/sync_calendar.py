"""Fetch next 2 days of Google Calendar events → data/calendar_events.json.

Uses scripts.google_auth (shared OAuth covering Calendar + Gmail).
Run on a schedule via run_calendar_sync.bat.

Optional env:
  GOOGLE_CALENDAR_ID   — defaults to 'primary'
  CALENDAR_DAYS_AHEAD  — defaults to 2
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

from googleapiclient.discovery import build

sys.path.insert(0, str(Path(__file__).parent))
from google_auth import get_credentials, load_env  # noqa: E402

BASE = Path(__file__).parent.parent
DATA_DIR = BASE / "data"


def _normalize(events: list[dict]) -> list[dict]:
    out = []
    for e in events:
        start = e.get("start") or {}
        end_dt = e.get("end") or {}
        all_day = bool(start.get("date")) and not start.get("dateTime")
        s = start.get("date") if all_day else start.get("dateTime")
        en = end_dt.get("date") if all_day else end_dt.get("dateTime")
        if not s:
            continue
        out.append({
            "id":       e.get("id"),
            "summary":  e.get("summary") or "(no title)",
            "start":    s,
            "end":      en,
            "all_day":  all_day,
            "location": e.get("location") or "",
            "link":     e.get("htmlLink") or "",
        })
    return out


def main() -> None:
    load_env()
    creds = get_credentials()
    cal_id = os.environ.get("GOOGLE_CALENDAR_ID", "primary")
    days = int(os.environ.get("CALENDAR_DAYS_AHEAD", "2"))

    service = build("calendar", "v3", credentials=creds, cache_discovery=False)
    start = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=days + 1)

    result = service.events().list(
        calendarId=cal_id,
        timeMin=start.isoformat(),
        timeMax=end.isoformat(),
        maxResults=100,
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    data = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "window_from":  start.isoformat(),
        "window_to":    end.isoformat(),
        "timezone":     "Europe/London",
        "source":       "google-calendar-api",
        "events":       _normalize(result.get("items") or []),
    }
    out_path = DATA_DIR / "calendar_events.json"
    out_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Wrote {len(data['events'])} events to {out_path}")


if __name__ == "__main__":
    main()
