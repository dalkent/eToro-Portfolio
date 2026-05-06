"""Fetch recent Gmail messages → data/emails.json.

Uses the same OAuth client as sync_calendar.py via scripts.google_auth.
Default query: unread inbox, last 48h, max 15 threads.

Optional env:
  GMAIL_QUERY     — Gmail search syntax. Default: 'is:unread in:inbox newer_than:2d'
  GMAIL_MAX       — max messages to fetch. Default 15.
"""
from __future__ import annotations

import base64
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from googleapiclient.discovery import build

sys.path.insert(0, str(Path(__file__).parent))
from google_auth import get_credentials, load_env  # noqa: E402

BASE = Path(__file__).parent.parent
DATA_DIR = BASE / "data"


def _header(headers: list[dict], name: str) -> str:
    name_l = name.lower()
    for h in headers or []:
        if (h.get("name") or "").lower() == name_l:
            return h.get("value") or ""
    return ""


def _strip_name(from_line: str) -> str:
    # "Jane Doe <jane@example.com>"  →  "Jane Doe"
    #  "jane@example.com"           →  "jane@example.com"
    if "<" in from_line and ">" in from_line:
        name = from_line.split("<")[0].strip().strip('"')
        return name or from_line.split("<")[1].rstrip(">")
    return from_line.strip()


def _parse_internal_date(ms: str | int | None) -> str:
    if not ms:
        return ""
    try:
        ts = int(ms) / 1000
    except (TypeError, ValueError):
        return ""
    return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat(timespec="seconds")


def main() -> None:
    load_env()
    creds = get_credentials()
    query = os.environ.get("GMAIL_QUERY", "is:unread in:inbox newer_than:2d")
    max_msgs = int(os.environ.get("GMAIL_MAX", "15"))

    service = build("gmail", "v1", credentials=creds, cache_discovery=False)

    # Fetch IDs first — cheap.
    resp = service.users().messages().list(
        userId="me", q=query, maxResults=max_msgs,
    ).execute()
    msg_ids = [m["id"] for m in (resp.get("messages") or [])]

    # Hydrate each with metadata (cheap: no body fetch).
    out: list[dict] = []
    for mid in msg_ids:
        try:
            m = service.users().messages().get(
                userId="me", id=mid, format="metadata",
                metadataHeaders=["From", "Subject", "Date"],
            ).execute()
        except Exception:
            continue
        payload = m.get("payload") or {}
        headers = payload.get("headers") or []
        out.append({
            "id":       m.get("id"),
            "thread":   m.get("threadId"),
            "from_raw": _header(headers, "From"),
            "from":     _strip_name(_header(headers, "From")),
            "subject":  _header(headers, "Subject") or "(no subject)",
            "snippet":  (m.get("snippet") or "").strip(),
            "received": _parse_internal_date(m.get("internalDate")),
            "unread":   "UNREAD" in (m.get("labelIds") or []),
            "link":     f"https://mail.google.com/mail/u/0/#inbox/{m.get('threadId')}",
        })

    # Total unread count for the badge (separate call — fast).
    unread_total = 0
    try:
        lbl = service.users().labels().get(userId="me", id="INBOX").execute()
        unread_total = int(lbl.get("messagesUnread") or 0)
    except Exception:
        pass

    data = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "query":        query,
        "unread_total": unread_total,
        "messages":     out,
    }
    out_path = DATA_DIR / "emails.json"
    out_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Wrote {len(out)} messages (total unread {unread_total}) to {out_path}")


if __name__ == "__main__":
    main()
