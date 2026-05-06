#!/usr/bin/env python3
"""
sync_t212.py
────────────
Fetches your Trading 212 ISA portfolio via the public API and writes
data/t212_portfolio.json for the combined dashboard.

Requirements (env vars, loaded from t212.env by run_combined.py):
    T212_API_KEY    — API key from the T212 app (Settings → API (Beta))
    T212_HOST       — https://live.trading212.com  (or demo.trading212.com)

Endpoints used:
    GET /api/v0/equity/account/info   — account currency
    GET /api/v0/equity/account/cash   — free / invested / total / ppl
    GET /api/v0/equity/portfolio      — open positions

Output:
    data/t212_portfolio.json
"""

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import requests

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"
OUT_FILE = DATA_DIR / "t212_portfolio.json"
INSTRUMENTS_CACHE = DATA_DIR / "t212_instruments.json"
LOG_FILE = LOGS_DIR / "sync_t212.log"

DATA_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

API_KEY    = os.getenv("T212_API_KEY")
API_KEY_ID = os.getenv("T212_API_KEY_ID")          # only used for T212_AUTH_MODE=basic
HOST       = os.getenv("T212_HOST", "https://live.trading212.com").rstrip("/")
AUTH_MODE  = (os.getenv("T212_AUTH_MODE") or "raw").lower()   # "raw" or "basic"


def auth_header() -> str:
    """Return the value for the Authorization header per T212_AUTH_MODE."""
    if AUTH_MODE == "basic":
        import base64
        if not API_KEY_ID:
            raise RuntimeError(
                "T212_AUTH_MODE=basic requires T212_API_KEY_ID (your API Key ID from the app)"
            )
        token = base64.b64encode(f"{API_KEY_ID}:{API_KEY}".encode()).decode()
        return f"Basic {token}"
    # default: raw key as Authorization value (current SDK behaviour)
    return API_KEY


def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    print(line)


def get(path: str, retries: int = 3):
    """GET with T212 auth, basic 429 back-off, surfaces response body on errors."""
    url = f"{HOST}/api/v0{path}"
    headers = {"Authorization": auth_header(), "Accept": "application/json"}
    for attempt in range(retries):
        r = requests.get(url, headers=headers, timeout=20)
        if r.status_code == 429:
            wait = 2 ** attempt
            log(f"  rate-limited on {path}, sleeping {wait}s")
            time.sleep(wait)
            continue
        if not r.ok:
            body = (r.text or "").strip()[:400]
            log(f"  HTTP {r.status_code} on {path} — body: {body!r}")
            if r.status_code == 401:
                log("  → 401 means the key was rejected. Common causes:")
                log("     1. Key pasted into the wrong file (must be t212.env, not t212.env.example)")
                log("     2. Key generated on Practice but T212_HOST points at live (or vice-versa)")
                log("        Practice host: https://demo.trading212.com")
                log("     3. Key lacks required scopes — enable 'View portfolio' + 'Account data'")
                log("     4. Key was rotated/revoked")
            r.raise_for_status()
        return r.json()
    raise RuntimeError(f"T212 API gave up after {retries} tries: {path}")


def paginate(path: str, params: dict | None = None, max_pages: int = 200) -> list:
    """
    Fetch all pages from a T212 history endpoint.

    T212 returns a ready-to-use `nextPagePath` on each page — we follow it
    verbatim rather than trying to reconstruct query params ourselves (their
    cursor is paired with a `time` param and they reject requests that only
    send the cursor).

    On any error during pagination, returns the items collected so far
    rather than losing them.
    """
    items = []
    qp = dict(params or {})
    qp["limit"] = 50
    qs = "&".join(f"{k}={v}" for k, v in qp.items())
    next_path = f"{path}?{qs}" if qs else path

    for _ in range(max_pages):
        try:
            data = get(next_path)
        except Exception as e:
            log(f"  pagination stopped on {next_path} — {e}; returning {len(items)} items so far")
            break
        batch = data.get("items") if isinstance(data, dict) else data
        if not batch:
            break
        items.extend(batch)

        npp = data.get("nextPagePath") if isinstance(data, dict) else None
        if not npp:
            break

        # T212 returns nextPagePath in several shapes. Normalise them all back
        # to `path + "?" + query` so we always target the right endpoint.
        npp = npp.lstrip("?")
        if npp.startswith("/api/v0"):
            npp = npp[len("/api/v0"):]
        if npp.startswith("/"):
            # full endpoint path with its own querystring — use as-is
            next_path = npp
        else:
            # bare query string — stick it back onto the original endpoint
            next_path = f"{path}?{npp}"
    return items


def load_instruments(force_refresh: bool = False) -> dict:
    """
    Returns {raw_ticker: {"name": str, "currency": str, "type": str}} for every
    T212 instrument. Cached to disk because the list has thousands of entries
    and rarely changes. Refreshes automatically if cache is >7 days old.
    """
    from datetime import timedelta, datetime as dt
    if INSTRUMENTS_CACHE.exists() and not force_refresh:
        age = dt.now() - dt.fromtimestamp(INSTRUMENTS_CACHE.stat().st_mtime)
        if age < timedelta(days=7):
            try:
                cached = json.loads(INSTRUMENTS_CACHE.read_text(encoding="utf-8"))
                log(f"  instruments: cache hit ({len(cached)} entries, age {age.days}d)")
                return cached
            except Exception:
                pass

    log("  instruments: fetching full metadata list (this is large) ...")
    try:
        raw = get("/equity/metadata/instruments")
    except Exception as e:
        log(f"  WARN: could not fetch instruments — {e}")
        return {}
    out = {
        entry["ticker"]: {
            "name":     entry.get("name") or entry.get("shortName") or "",
            "currency": entry.get("currencyCode", ""),
            "type":     entry.get("type", ""),
        }
        for entry in raw if entry.get("ticker")
    }
    INSTRUMENTS_CACHE.write_text(json.dumps(out), encoding="utf-8")
    log(f"  instruments: cached {len(out)} entries")
    return out


def clean_ticker(t212_ticker: str) -> str:
    """
    T212 tickers look like 'SHELl_EQ', 'AAPL_US_EQ', 'VUSA_EQ', 'BTC_USD'.
    Strip the '_EQ' / '_US_EQ' suffix for display. Leave the rest alone.
    """
    t = t212_ticker
    if t.endswith("_US_EQ"):
        return t[:-6]
    if t.endswith("_EQ"):
        return t[:-3]
    return t


def main():
    if not API_KEY:
        log("ERROR: T212_API_KEY not set. Create t212.env from t212.env.example.")
        sys.exit(1)

    log(f"T212 host: {HOST}  auth mode: {AUTH_MODE}")

    info = get("/equity/account/info")
    account_currency = info.get("currencyCode", "GBP")
    log(f"  account currency: {account_currency}  id: {info.get('id')}")

    cash = get("/equity/account/cash")
    log(f"  cash: free={cash.get('free')} invested={cash.get('invested')} "
        f"total={cash.get('total')} ppl={cash.get('ppl')} "
        f"result={cash.get('result')} pieCash={cash.get('pieCash')}")

    positions = get("/equity/portfolio")
    log(f"  positions: {len(positions)}")

    instruments = load_instruments()

    # ── History: transactions (deposits/withdrawals) + dividends ─────────────
    # Requires "History" scope on the key. Tolerant to missing scope — just
    # skip and log if we get a 4xx.
    transactions, dividends = [], []
    try:
        transactions = paginate("/history/transactions")
        log(f"  transactions fetched: {len(transactions)}")
    except Exception as e:
        log(f"  WARN: could not fetch /history/transactions — {e}")

    try:
        dividends = paginate("/history/dividends")
        log(f"  dividends fetched: {len(dividends)}")
    except Exception as e:
        log(f"  WARN: could not fetch /history/dividends — {e}")

    # Break down by transaction type so we can see exactly what's there
    type_counts = {}
    type_sums   = {}
    for t in transactions:
        t_type = (t.get("type") or "UNKNOWN").upper()
        amount = float(t.get("amount") or 0)
        type_counts[t_type] = type_counts.get(t_type, 0) + 1
        type_sums[t_type]   = type_sums.get(t_type, 0.0) + amount
    log("  transaction type breakdown:")
    for t_type in sorted(type_counts.keys()):
        log(f"     {t_type:<22} count={type_counts[t_type]:>4}  "
            f"sum={type_sums[t_type]:+,.2f}")

    # Money-in types: standard DEPOSIT plus TRANSFER (inbound ISA transfers from
    # other providers arrive as TRANSFER and are still the user's own money).
    DEPOSIT_TYPES    = {"DEPOSIT", "TRANSFER"}
    WITHDRAWAL_TYPES = {"WITHDRAW", "WITHDRAWAL"}
    deposits_total    = sum(abs(type_sums.get(k, 0)) for k in DEPOSIT_TYPES
                            if (type_sums.get(k, 0) or 0) > 0)
    withdrawals_total = sum(abs(type_sums.get(k, 0)) for k in WITHDRAWAL_TYPES)
    # Anything else that added to the balance (interest, promos, referrals).
    bonus_credits = 0.0
    for t_type, s in type_sums.items():
        if t_type in DEPOSIT_TYPES or t_type in WITHDRAWAL_TYPES:
            continue
        if s > 0:
            bonus_credits += s

    dividends_total = sum(float(d.get("amount") or 0) for d in dividends)
    log(f"  lifetime: deposits £{deposits_total:,.2f}  "
        f"withdrawals £{withdrawals_total:,.2f}  "
        f"dividends £{dividends_total:,.2f}  "
        f"bonus credits £{bonus_credits:,.2f}")

    out = {
        "generated_at":     datetime.now().isoformat(timespec="seconds"),
        "account_currency": account_currency,
        "account_id":       info.get("id"),
        "cash": {
            "free":     cash.get("free", 0),
            "invested": cash.get("invested", 0),
            "total":    cash.get("total", 0),
            "ppl":      cash.get("ppl", 0),
            "result":   cash.get("result", 0),
            "pie_cash": cash.get("pieCash", 0),
        },
        "positions": [
            {
                "ticker":           clean_ticker(p["ticker"]),
                "raw_ticker":       p["ticker"],
                "name":             (instruments.get(p["ticker"], {}) or {}).get("name", ""),
                "instrument_ccy":   (instruments.get(p["ticker"], {}) or {}).get("currency", ""),
                "quantity":         p.get("quantity", 0),
                "average_price":    p.get("averagePrice", 0),   # instrument ccy
                "current_price":    p.get("currentPrice", 0),   # instrument ccy
                "ppl":              p.get("ppl", 0),            # account ccy
                "fx_ppl":           p.get("fxPpl"),
                "initial_fill":     p.get("initialFillDate"),
                "pie_quantity":     p.get("pieQuantity", 0),
            }
            for p in positions
        ],
        "lifetime": {
            "deposits":          deposits_total,
            "withdrawals":       withdrawals_total,
            "bonus_credits":     bonus_credits,
            "net_deposited":     deposits_total - withdrawals_total,
            "dividends":         dividends_total,
            "transaction_count": len(transactions),
            "dividend_count":    len(dividends),
            "type_breakdown":    {k: {"count": type_counts[k], "sum": type_sums[k]}
                                  for k in sorted(type_counts)},
        },
    }

    OUT_FILE.write_text(json.dumps(out, indent=2), encoding="utf-8")
    log(f"  wrote {OUT_FILE}")


if __name__ == "__main__":
    main()
