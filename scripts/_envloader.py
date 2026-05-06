"""Shared env loader.

Reads `etoro.env` from the project root (parent of this scripts/ folder) and
populates os.environ for any vars not already set in the shell. Imported for
side effects only:

    from _envloader import load  # noqa: F401
    load()

Or just::

    import _envloader  # noqa: F401  (loads on import)
"""
from __future__ import annotations
import os
from pathlib import Path

_LOADED = False


def load(env_path: str | Path | None = None) -> None:
    """Idempotent. Defaults to <project_root>/etoro.env."""
    global _LOADED
    if _LOADED:
        return
    if env_path is None:
        env_path = Path(__file__).resolve().parent.parent / "etoro.env"
    env_path = Path(env_path)
    if not env_path.exists():
        _LOADED = True
        return
    # Read with utf-8-sig so a BOM doesn't break the first key
    for line in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        # Strip optional surrounding quotes
        value = value.strip().strip('"').strip("'")
        # setdefault: shell-set vars win
        os.environ.setdefault(key, value)
    _LOADED = True


# Auto-load on import for the common case
load()
