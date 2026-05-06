"""Platform-aware path resolution.

Replaces hardcoded C:\\Users\\Neil\\... references throughout the codebase.
Detects OS at import time. Override any path via environment variable.
"""
from __future__ import annotations
import os
import sys
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent.parent

def _resolve_drive() -> Path:
    if os.environ.get("DRIVE_DIR"):
        return Path(os.environ["DRIVE_DIR"])
    if sys.platform == "win32":
        return Path(os.path.expanduser("~")) / "My Drive"
    home = Path.home()
    candidates = [
        home / "Library" / "CloudStorage" / "GoogleDrive-ndaley1313@gmail.com" / "My Drive",
        home / "Google Drive" / "My Drive",
        home / "Google Drive",
    ]
    for c in candidates:
        if c.exists():
            return c
    return candidates[0]

DRIVE_DIR  = _resolve_drive()
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR",  str(DRIVE_DIR / "Upload")))
SYNC_DIR   = Path(os.environ.get("SYNC_DIR",    str(DRIVE_DIR / "eToro Sync")))
VAULT_DIR  = Path(os.environ.get("VAULT_ROOT",  str(DRIVE_DIR / "Daley's Brain")))

DATA_DIR       = REPO_DIR / "data"
DASHBOARDS_DIR = REPO_DIR / "dashboards"
SCRIPTS_DIR    = REPO_DIR / "scripts"

def diagnostic():
    print(f"REPO_DIR       = {REPO_DIR}")
    print(f"DRIVE_DIR      = {DRIVE_DIR}     exists={DRIVE_DIR.exists()}")
    print(f"UPLOAD_DIR     = {UPLOAD_DIR}    exists={UPLOAD_DIR.exists()}")
    print(f"SYNC_DIR       = {SYNC_DIR}      exists={SYNC_DIR.exists()}")
    print(f"VAULT_DIR      = {VAULT_DIR}     exists={VAULT_DIR.exists()}")
    print(f"DATA_DIR       = {DATA_DIR}      exists={DATA_DIR.exists()}")
    print(f"DASHBOARDS_DIR = {DASHBOARDS_DIR} exists={DASHBOARDS_DIR.exists()}")

if __name__ == "__main__":
    diagnostic()
