"""refresh_finances.py — One-shot wrapper: sync Google Sheet + regenerate HTML.

Designed to be invoked by Task Scheduler via `pythonw.exe` so no console
window appears. Logs to logs/finances_refresh.log.
"""
from __future__ import annotations

import sys
import traceback
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent.parent
LOG = BASE / "logs" / "finances_refresh.log"
LOG.parent.mkdir(exist_ok=True)

sys.path.insert(0, str(Path(__file__).parent))


def log(msg: str) -> None:
    line = f"[{datetime.now().isoformat(timespec='seconds')}] {msg}\n"
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line)


def main() -> int:
    try:
        log("=== refresh_finances start ===")
        import sync_finances  # noqa: E402
        sync_finances.main()
        log("sync_finances done")
        import generate_finances_dashboard  # noqa: E402
        generate_finances_dashboard.main()
        log("generate_finances_dashboard done")
        log("=== refresh_finances complete ===")
        return 0
    except Exception as e:
        log(f"ERROR: {e}")
        log(traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
