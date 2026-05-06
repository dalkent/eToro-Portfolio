"""Shared Google OAuth helper used by sync_calendar.py and sync_gmail.py.

Single OAuth client JSON authorises both Calendar (read-only) and Gmail
(read-only) so the user only ever does the browser consent flow once.

If the existing token covers fewer scopes than requested, it is discarded and
a new consent flow starts.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    sys.exit(
        "Missing packages. Install with:\n"
        "  pip install google-api-python-client google-auth-oauthlib"
    )

BASE = Path(__file__).parent.parent
CREDS_DIR = BASE / "credentials"
CREDS_DIR.mkdir(exist_ok=True)

SCOPES = [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]
TOKEN_PATH = CREDS_DIR / "google_token.json"
CLIENT_SECRET_PATH = CREDS_DIR / "google_calendar_client.json"


def _covers_scopes(creds: Credentials) -> bool:
    have = set(creds.scopes or [])
    need = set(SCOPES)
    return need.issubset(have)


def get_credentials() -> Credentials:
    creds: Credentials | None = None
    if TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except Exception:
            creds = None
    # If token is missing scopes, drop it so we re-auth with the combined list.
    if creds and not _covers_scopes(creds):
        creds = None
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
            return creds
        except Exception:
            creds = None
    if not CLIENT_SECRET_PATH.exists():
        sys.exit(
            f"Missing OAuth client file at {CLIENT_SECRET_PATH}.\n"
            "Download from Google Cloud Console → APIs & Services → Credentials."
        )
    flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET_PATH), SCOPES)
    creds = flow.run_local_server(port=0)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    return creds


def load_env() -> None:
    env_path = BASE / "etoro.env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        os.environ.setdefault(k.strip(), v.strip())
