#!/usr/bin/env python3
"""
sync_macro.py
─────────────
Pulls global macro data + markets + FX and writes data/macro.json for the
macro dashboard.

Data sources:
  • FRED  — Federal Reserve Economic Data (GDP, CPI, unemployment,
            central bank rates, long yields for US / UK / Euro / Japan / China).
            Free API key: https://fred.stlouisfed.org/docs/api/api_key.html
  • yfinance — stock indices, FX, live 10Y yields as sanity check.

Env:
  FRED_API_KEY   — required. Add to etoro.env as: FRED_API_KEY=...

Output:
  data/macro.json
"""

import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

# Load etoro.env from project root before reading os.getenv anywhere below.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import _envloader  # noqa: F401  (side-effect import: populates os.environ)

import requests

BASE_DIR  = Path(__file__).parent.parent
DATA_DIR  = BASE_DIR / "data"
LOGS_DIR  = BASE_DIR / "logs"
OUT_FILE  = DATA_DIR / "macro.json"
LOG_FILE  = LOGS_DIR / "sync_macro.log"

DATA_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

FRED_API_KEY = os.getenv("FRED_API_KEY", "").strip()
FRED_BASE    = "https://api.stlouisfed.org/fred"

# ── Which FRED series to pull ───────────────────────────────────────────────
# Each entry: key → (series_id, friendly label, kind)
# kind = "rate"     — series already reports a %, pass through as-is
#      = "level_yoy"— series is an index/level; we compute YoY from observations
#      = "level_qoq"— series is quarterly level; we compute QoQ
FRED_SERIES = {
    # US
    "us_cpi_yoy":       ("CPIAUCSL",         "US CPI (YoY)",           "level_yoy"),
    "us_gdp_yoy":       ("GDPC1",            "US Real GDP (YoY)",      "level_yoy"),
    "us_unemployment":  ("UNRATE",           "US Unemployment",        "rate"),
    "us_fed_funds":     ("DFF",              "Fed Funds Rate",         "rate"),
    "us_10y":           ("DGS10",            "US 10Y Treasury",        "rate"),

    # UK
    "uk_cpi_yoy":       ("CPALTT01GBM659N",  "UK CPI (YoY)",           "rate"),    # MEI YoY series
    "uk_gdp_yoy":       ("NGDPRSAXDCGBQ",    "UK Real GDP (YoY)",      "level_yoy"),
    "uk_unemployment":  ("LRHUTTTTGBM156S",  "UK Unemployment",        "rate"),
    "uk_bank_rate":     ("IR3TIB01GBM156N",  "UK 3M interbank",        "rate"),
    "uk_10y":           ("IRLTLT01GBM156N",  "UK 10Y Gilt",            "rate"),

    # Eurozone
    "ez_cpi_yoy":       ("CP0000EZ19M086NEST","Euro area HICP (YoY)",  "level_yoy"),
    "ez_gdp_yoy":       ("CLVMNACSCAB1GQEA19","Euro area Real GDP (YoY)","level_yoy"),
    "ez_unemployment":  ("LRHUTTTTEZM156S",  "Euro area Unemployment", "rate"),
    "ez_rate":          ("ECBDFR",           "ECB Deposit Rate",       "rate"),
    "de_10y":           ("IRLTLT01DEM156N",  "German 10Y Bund",        "rate"),

    # Japan — CPALTT01JPM659N reports YoY directly (rate series), always fresher than JPNCPIALLMINMEI
    "jp_cpi_yoy":       ("CPALTT01JPM659N",  "Japan CPI (YoY)",        "rate"),
    "jp_unemployment":  ("LRHUTTTTJPM156S",  "Japan Unemployment",     "rate"),
    "jp_10y":           ("IRLTLT01JPM156N",  "Japan 10Y JGB",          "rate"),

    # China
    "cn_cpi_yoy":       ("CPALTT01CNM659N",  "China CPI (YoY)",        "rate"),
}

# ── Yahoo tickers for markets + FX ──────────────────────────────────────────
YF_MARKETS = {
    # Equity indices — key: (ticker, friendly label, region)
    "ftse100":       ("^FTSE",       "FTSE 100",        "UK"),
    "ftse250":       ("^FTMC",       "FTSE 250",        "UK"),
    "sp500":         ("^GSPC",       "S&P 500",         "US"),
    "nasdaq":        ("^IXIC",       "Nasdaq Composite","US"),
    "dow":           ("^DJI",        "Dow Jones",       "US"),
    "eurostoxx":     ("^STOXX50E",   "Euro Stoxx 50",   "Eurozone"),
    "dax":           ("^GDAXI",      "DAX",             "Germany"),
    "cac":           ("^FCHI",       "CAC 40",          "France"),
    "nikkei":        ("^N225",       "Nikkei 225",      "Japan"),
    "hangseng":      ("^HSI",        "Hang Seng",       "Hong Kong"),
    "shanghai":      ("000001.SS",   "Shanghai Comp.",  "China"),
    "vix":           ("^VIX",        "VIX",             "Global"),
}

YF_FX = {
    "gbpusd":  ("GBPUSD=X", "GBP/USD"),
    "gbpeur":  ("GBPEUR=X", "GBP/EUR"),
    "gbpjpy":  ("GBPJPY=X", "GBP/JPY"),
    "gbptwd":  ("GBPTWD=X", "GBP/TWD"),
    "eurusd":  ("EURUSD=X", "EUR/USD"),
    "usdjpy":  ("USDJPY=X", "USD/JPY"),
    "usdcny":  ("USDCNY=X", "USD/CNY"),
    "dxy":     ("DX-Y.NYB", "DXY (USD index)"),
}


# ── Live sovereign yields ───────────────────────────────────────────────────
# US 10y comes from yfinance (^TNX = quoted in tenths of a percent, so divide by 10).
# UK 10y comes from BoE's IADB CSV (Yahoo doesn't carry the UK 10y reliably).
# Each entry maps to a key the dashboard uses to compare against the FRED
# monthly average + the model assumption.
LIVE_YIELDS_YF = {
    "us_10y_live": ("^TNX",  "US 10Y Treasury (live)", 1),  # last factor: divide raw by this (0.1 -> result in % directly: 43 -> 4.3)
    "us_30y_live": ("^TYX",  "US 30Y Treasury (live)", 1),
}

# Bank of England IADB series codes (daily, business-day). All values are %.
LIVE_YIELDS_BOE = {
    "uk_10y_live": ("IUDMNZC", "UK 10Y Gilt (live)"),
    # "uk_2y_live":  ("IUDSNB2", "UK 2Y Gilt (live)"),  # disabled - need correct BoE code
}


# ── ONS (UK Office for National Statistics) — direct, always current ───────
# Free, no API key. ONS retired their old api.ons.gov.uk endpoint; the
# canonical site URL with /data appended now serves the timeseries JSON.
# Each tuple is (category_path, dataset_code, series_code, label, kind).
# category_path is the URL slug under www.ons.gov.uk where the series lives.
# uk_bank_rate (OOGA) is not in ONS - it's a BoE figure - so it's omitted.
ONS_SERIES = {
    "uk_cpi_yoy":       ("economy/inflationandpriceindices",            "MM23", "D7G7", "UK CPI (YoY)",      "pct_yoy"),
    "uk_cpih_yoy":      ("economy/inflationandpriceindices",            "MM23", "L55O", "UK CPIH (YoY)",     "pct_yoy"),
    "uk_gdp_yoy":       ("economy/grossdomesticproductgdp",             "QNA",  "IHYQ", "UK Real GDP (YoY)", "pct_yoy"),
    "uk_gdp_qoq":       ("economy/grossdomesticproductgdp",             "QNA",  "IHYR", "UK Real GDP (QoQ)", "pct_qoq"),
    "uk_unemployment":  ("employmentandlabourmarket/peoplenotinwork/unemployment", "LMS", "MGSX", "UK Unemployment", "pct"),
}
ONS_BASE = "https://www.ons.gov.uk"


# Map ONS month abbreviations to ISO month numbers, e.g. "2026 MAR" -> "2026-03-01"
_ONS_MONTHS = {
    "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
    "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
}


def _ons_normalise_date(raw_date: str, year: str, month_full: str, quarter: str) -> str:
    """Convert ONS date strings to ISO YYYY-MM-DD so they compare cleanly with FRED dates.
    Fall back to the raw value if we can't parse."""
    if raw_date and len(raw_date) == 8 and " " in raw_date:
        # "2026 MAR" -> "2026-03-01"
        parts = raw_date.split()
        if len(parts) == 2 and parts[1] in _ONS_MONTHS:
            return f"{parts[0]}-{_ONS_MONTHS[parts[1]]}-01"
    if year and quarter:
        # "2025 Q4" -> "2025-12-01" (use the last month of the quarter)
        q = quarter.replace("Q", "").strip()
        if q in ("1", "2", "3", "4"):
            month = {"1": "03", "2": "06", "3": "09", "4": "12"}[q]
            return f"{year}-{month}-01"
    if year and month_full:
        # Fallback if month is given as full name
        for abbr, num in _ONS_MONTHS.items():
            if month_full.upper().startswith(abbr):
                return f"{year}-{num}-01"
    if year and not quarter and not month_full:
        return f"{year}-01-01"   # year-only series
    return raw_date or ""


def ons_latest(category: str, dataset: str, series: str) -> dict | None:
    """Fetch the most recent observation for an ONS timeseries."""
    url = f"{ONS_BASE}/{category}/timeseries/{series.lower()}/{dataset.lower()}/data"
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        log(f"  ONS {dataset}/{series}: error {e}")
        return None

    # ONS returns observation blocks: 'years', 'quarters', 'months'. Pick the
    # finest-grained pool that has data so we don't lose a recent monthly print
    # to an older yearly aggregate.
    pools = [data.get("months", []), data.get("quarters", []), data.get("years", [])]
    for obs in pools:
        if not obs:
            continue
        latest = obs[-1]
        prior  = obs[-2] if len(obs) >= 2 else None
        try:
            val  = float(latest["value"])
            prev = float(prior["value"]) if prior else None
        except (KeyError, ValueError, TypeError):
            continue
        date = _ons_normalise_date(
            latest.get("date", ""),
            latest.get("year", ""),
            latest.get("month", ""),
            latest.get("quarter", ""),
        )
        prev_date = ""
        if prior:
            prev_date = _ons_normalise_date(
                prior.get("date", ""),
                prior.get("year", ""),
                prior.get("month", ""),
                prior.get("quarter", ""),
            )
        return {
            "value":     val,
            "prev":      prev,
            "date":      date,
            "prev_date": prev_date,
            "series_id": f"ONS {dataset}/{series}",
        }
    return None


def apply_ons_overrides(out: dict, *, log_fn=print) -> int:
    """Walk ONS_SERIES and override the FRED entry for each key when the ONS
    observation is fresher (or when no FRED entry exists at all).

    Why this exists: FRED's OECD-mirrored UK CPI series (CPALTT01GBM659N) has
    repeatedly gone stale by 6-12 months. Pulling direct from ONS keeps the
    indicators table current.
    """
    fred = out.setdefault("fred", {})
    applied = 0
    for key, (category, dataset, series, label, kind) in ONS_SERIES.items():
        ons = ons_latest(category, dataset, series)
        if not ons:
            continue
        existing = fred.get(key) or {}
        existing_date = existing.get("date") or ""
        ons_date = ons.get("date") or ""
        # Apply if no FRED entry, or ONS is at least as fresh
        if not existing_date or ons_date > existing_date:
            fred[key] = {
                **ons,
                "label": label,
                "kind":  kind,
            }
            applied += 1
            log_fn(f"    ONS override: {key} -> {ons['value']} ({ons_date}) [was {existing.get('value')} ({existing_date or 'none'})]")
    return applied


# ── Valuation-model assumptions (must match scripts/valuation.py) ───────────
# Surfacing these on the dashboard so you can see when reality has drifted.
MODEL_ASSUMPTIONS = {
    "rf_uk_pct":         4.9,    # UK 10Y gilt assumption
    "rf_us_pct":         4.5,    # US 10Y treasury assumption
    "erp_pct":           5.0,    # Equity risk premium
    "wacc_default_pct":  9.0,    # Default WACC
    "growth_5y_pct":     5.0,    # 5-year FCF/dividend growth
    "terminal_g_pct":    2.5,    # Perpetual growth
}


# ── Logging ─────────────────────────────────────────────────────────────────

def log(msg: str):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    print(line)


# ── FRED fetch ──────────────────────────────────────────────────────────────

def _fred_observations(series_id: str, limit: int = 15) -> list:
    """Fetch most-recent N observations, descending by date, as list of dicts."""
    if not FRED_API_KEY:
        return []
    params = {
        "series_id":  series_id,
        "api_key":    FRED_API_KEY,
        "file_type":  "json",
        "sort_order": "desc",
        "limit":      limit,
    }
    try:
        r = requests.get(f"{FRED_BASE}/series/observations",
                         params=params, timeout=15)
        r.raise_for_status()
    except Exception as e:
        log(f"  FRED {series_id}: error {e}")
        return []
    obs = r.json().get("observations", [])
    cleaned = []
    for o in obs:
        try:
            o = dict(o)
            o["_val"] = float(o["value"])
            cleaned.append(o)
        except (KeyError, ValueError):
            continue
    return cleaned


def fred_latest(series_id: str, kind: str = "rate") -> dict | None:
    """
    Fetch the latest observation and compute the right headline figure:
      kind="rate"       → pass-through (series already a %).
      kind="level_yoy"  → compute YoY from 13 months / 5 quarters ago.
      kind="level_qoq"  → compute QoQ from prior observation.
    """
    # We need enough history to compute YoY: 14 obs for monthly, 6 for quarterly.
    obs = _fred_observations(series_id, limit=20)
    if not obs:
        return None

    latest = obs[0]
    date   = latest["date"]
    raw    = latest["_val"]

    if kind == "rate":
        prior = obs[1] if len(obs) > 1 else None
        return {
            "value":     raw,
            "prev":      prior["_val"] if prior else None,
            "date":      date,
            "prev_date": prior["date"] if prior else None,
            "series_id": series_id,
        }

    # Determine frequency by looking at gap between first two dates
    def _parse_dt(s):
        try:
            return datetime.strptime(s, "%Y-%m-%d")
        except ValueError:
            return None

    # For YoY, find the observation ~365 days before latest
    latest_dt = _parse_dt(date)
    target_dt = latest_dt - timedelta(days=365) if latest_dt else None

    if target_dt:
        # Pick the obs closest to one-year-ago date
        best = None
        best_gap = None
        for o in obs[1:]:
            o_dt = _parse_dt(o["date"])
            if not o_dt:
                continue
            gap = abs((o_dt - target_dt).days)
            if best is None or gap < best_gap:
                best, best_gap = o, gap
        if best and best_gap is not None and best_gap < 45:  # within ~6 weeks of 1y
            yoy = (raw / best["_val"] - 1) * 100
            # Previous YoY: compare prior obs vs its year-ago
            prior = obs[1] if len(obs) > 1 else None
            prev_yoy = None
            if prior:
                prior_dt = _parse_dt(prior["date"])
                if prior_dt:
                    ptarget = prior_dt - timedelta(days=365)
                    pbest = None; pbest_gap = None
                    for o in obs[2:]:
                        o_dt = _parse_dt(o["date"])
                        if not o_dt:
                            continue
                        gap = abs((o_dt - ptarget).days)
                        if pbest is None or gap < pbest_gap:
                            pbest, pbest_gap = o, gap
                    if pbest and pbest_gap is not None and pbest_gap < 45:
                        prev_yoy = (prior["_val"] / pbest["_val"] - 1) * 100
            return {
                "value":     yoy,
                "prev":      prev_yoy,
                "date":      date,
                "prev_date": prior["date"] if prior else None,
                "series_id": series_id,
            }

    # Fallback if we don't have enough history
    return {
        "value":     None,
        "prev":      None,
        "date":      date,
        "series_id": series_id,
    }


# ── yfinance fetch ──────────────────────────────────────────────────────────

def yf_latest(ticker: str) -> dict | None:
    """Return latest close + prev close + pct change."""
    try:
        import yfinance as yf
    except ImportError:
        log("  yfinance not installed — markets + FX will be empty")
        return None

    try:
        hist = yf.Ticker(ticker).history(period="5d")
        if hist.empty:
            return None
        closes = hist["Close"].dropna()
        if len(closes) == 0:
            return None
        latest = float(closes.iloc[-1])
        prev   = float(closes.iloc[-2]) if len(closes) >= 2 else None
        change_pct = ((latest / prev) - 1) * 100 if prev else None
        return {
            "value":      latest,
            "prev":       prev,
            "change_pct": change_pct,
            "as_of":      hist.index[-1].strftime("%Y-%m-%d"),
            "ticker":     ticker,
        }
    except Exception as e:
        log(f"  yfinance {ticker}: error {e}")
        return None


# ── Live sovereign yield fetchers ───────────────────────────────────────────

def _yf_yield_latest(ticker: str, divide_by: float = 1.0) -> dict | None:
    """Live yield via yfinance. Some tickers (^TNX) report in tenths of a
    percent so the caller passes divide_by=0.1. Returns a normalised dict."""
    raw = yf_latest(ticker)
    if not raw:
        return None
    val = raw["value"] / divide_by if divide_by else raw["value"]
    return {
        "value":  val,
        "as_of":  raw["as_of"],
        "source": f"yfinance:{ticker}",
    }


def _boe_yield_latest(series_code: str) -> dict | None:
    """Live yield via Bank of England IADB CSV. Daily business-day series.
    Returns a dict with the latest available observation (4dp), or None.

    BoE blocks default User-Agents so we send a browser UA.
    """
    from datetime import date, timedelta
    end = date.today()
    start = end - timedelta(days=14)   # window enough to cover bank holidays + weekends
    url = "https://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp"
    params = {
        "csv.x":      "yes",
        "Datefrom":   start.strftime("%d/%b/%Y"),
        "Dateto":     end.strftime("%d/%b/%Y"),
        "SeriesCodes": series_code,
        "CSVF":       "TN",
        "UsingCodes": "Y",
        "VPD": "Y", "VFD": "N",
    }
    headers = {"User-Agent": "Mozilla/5.0 (compatible; FTSE-Valuation-Insights/1.0)"}
    try:
        r = requests.get(url, params=params, headers=headers, timeout=15)
        r.raise_for_status()
        # Response is CSV: "DATE,IUDMNZC\n14 Apr 2026,4.8307\n..."
        lines = [ln.strip() for ln in r.text.splitlines() if ln.strip()]
        if len(lines) < 2:
            log(f"  BoE {series_code}: no rows returned")
            return None
        # Walk from the bottom up to find the most recent row with a valid value
        for line in reversed(lines[1:]):
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 2:
                continue
            try:
                val = float(parts[1])
            except ValueError:
                continue
            # Convert "24 Apr 2026" -> "2026-04-24"
            try:
                d = datetime.strptime(parts[0], "%d %b %Y").strftime("%Y-%m-%d")
            except ValueError:
                d = parts[0]
            return {"value": val, "as_of": d, "source": f"BoE IADB {series_code}"}
        return None
    except Exception as e:
        log(f"  BoE {series_code}: error {e}")
        return None


# ── Main ────────────────────────────────────────────────────────────────────

FINNHUB_KEY = os.getenv("FINNHUB_API_KEY", "").strip()


def _finnhub_events(days_ahead: int = 21) -> list:
    """Pull upcoming economic releases from Finnhub (same source investing.com uses).

    Returns a list of {date, time, country, event, impact, actual, estimate, prev, unit}.
    Filtered to GB/US/EU and sorted by time ascending.
    """
    if not FINNHUB_KEY:
        return []
    from urllib.request import urlopen
    from urllib.parse import urlencode
    today = datetime.now().date()
    end = today + timedelta(days=days_ahead)
    url = (
        "https://finnhub.io/api/v1/calendar/economic?"
        + urlencode({"from": today.isoformat(), "to": end.isoformat(), "token": FINNHUB_KEY})
    )
    try:
        with urlopen(url, timeout=15) as r:
            data = json.loads(r.read())
    except Exception as e:
        log(f"  WARN: Finnhub events fetch failed: {e}")
        return []
    raw = (data.get("economicCalendar") or []) if isinstance(data, dict) else []
    keep_countries = {"GB", "UK", "US", "EU", "DE", "FR"}
    out = []
    for ev in raw:
        country = ev.get("country") or ""
        if country and country not in keep_countries:
            continue
        out.append({
            "time":     (ev.get("time") or "").replace("T", " ")[:16],
            "country":  country,
            "event":    ev.get("event") or "",
            "impact":   (ev.get("impact") or "").lower(),
            "actual":   ev.get("actual"),
            "estimate": ev.get("estimate"),
            "prev":     ev.get("prev"),
            "unit":     ev.get("unit") or "",
        })
    out.sort(key=lambda e: e.get("time") or "")
    return out


RELEASE_OVERRIDES = {
    # (country, event_name_lower_contains) -> fred_key
    ("GB", "unemployment rate"):             "uk_unemployment",
    ("US", "unemployment rate"):             "us_unemployment",
    ("EU", "unemployment rate"):             "ez_unemployment",
    ("DE", "unemployment rate"):             "de_unemployment",
    ("JP", "unemployment rate"):             "jp_unemployment",
    ("CN", "unemployment rate"):             "cn_unemployment",

    ("GB", "inflation rate yoy"):            "uk_cpi_yoy",
    ("US", "cpi yoy"):                       "us_cpi_yoy",
    ("US", "inflation rate yoy"):            "us_cpi_yoy",
    ("EU", "inflation rate yoy"):            "ez_cpi_yoy",
    ("EU", "core inflation rate yoy"):       "ez_cpi_yoy",
    ("JP", "inflation rate yoy"):            "jp_cpi_yoy",
    ("CN", "inflation rate yoy"):            "cn_cpi_yoy",

    ("GB", "gdp yoy"):                       "uk_gdp_yoy",
    ("GB", "gdp growth rate yoy"):           "uk_gdp_yoy",
    ("US", "gdp yoy"):                       "us_gdp_yoy",
    ("US", "gdp growth rate yoy"):           "us_gdp_yoy",
    ("EU", "gdp yoy"):                       "ez_gdp_yoy",

    ("GB", "boe interest rate decision"):    "uk_bank_rate",
    ("US", "fed interest rate decision"):    "us_fed_funds",
    ("US", "fomc interest rate decision"):   "us_fed_funds",
    ("EU", "ecb interest rate decision"):    "ez_rate",
    ("EU", "ecb deposit facility rate"):     "ez_rate",
}


def _apply_release_overrides(out: dict) -> int:
    events = out.get("events") or []
    fred = out.setdefault("fred", {})
    applied = 0
    for ev in events:
        actual = ev.get("actual")
        if actual in (None, ""):
            continue
        name = (ev.get("event") or "").lower()
        country = ev.get("country") or ""
        # First try exact mapping, else substring match
        fred_key = None
        for (c, keyword), fk in RELEASE_OVERRIDES.items():
            if c == country and keyword in name:
                fred_key = fk
                break
        if not fred_key:
            continue
        ev_date = (ev.get("time") or "")[:10]
        existing = fred.get(fred_key) or {}
        existing_date = existing.get("date") or ""
        if existing_date and existing_date >= ev_date:
            continue  # our FRED value is already at least as fresh
        try:
            actual_f = float(actual)
        except (TypeError, ValueError):
            continue
        fred[fred_key] = {
            "value":     actual_f,
            "prev":      existing.get("value"),
            "date":      ev_date,
            "prev_date": existing.get("date"),
            "label":     existing.get("label") or f"{country} {name.title()}",
            "kind":      existing.get("kind", "rate"),
            "series_id": f"finnhub: {ev.get('event')}",
        }
        applied += 1
    return applied


def main():
    log(f"Fetching macro data ...")
    out = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "fred":         {},
        "markets":      {},
        "fx":           {},
        "yields_live":  {},
        "assumptions":  MODEL_ASSUMPTIONS,
    }

    # FRED series
    if not FRED_API_KEY:
        log("  WARN: FRED_API_KEY not set — macro series will be empty.")
        log("        Add FRED_API_KEY=... to etoro.env (free key from")
        log("        https://fred.stlouisfed.org/docs/api/api_key.html)")
    else:
        for key, (series_id, label, kind) in FRED_SERIES.items():
            data = fred_latest(series_id, kind=kind)
            if data:
                data["label"] = label
                data["kind"]  = kind
                out["fred"][key] = data
        log(f"  FRED: fetched {len(out['fred'])} / {len(FRED_SERIES)} series")

    # ── ONS overrides for UK series (auto-fresh from ons.gov.uk) ───────────
    # Some FRED-mirrored OECD series (UK CPI, GDP, etc.) lag the underlying ONS
    # data by months. Pull direct from ONS and override the FRED entry when ONS
    # is fresher. This runs before manual CSV so manual values still take
    # precedence if you ever need to override even ONS.
    ons_count = apply_ons_overrides(out, log_fn=log)
    if ons_count:
        log(f"  ONS overrides applied: {ons_count}")

    # ── Manual overrides from data/macro_manual.csv ────────────────────────
    # These take precedence over FRED and ONS — use them when both are stale.
    manual_path = DATA_DIR / "macro_manual.csv"
    manual_count = 0
    if manual_path.exists():
        import csv
        with open(manual_path, encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or line.lower().startswith("key,"):
                    continue
                parts = [p.strip() for p in line.split(",")]
                if len(parts) < 3:
                    continue
                key, val_s, date = parts[0], parts[1], parts[2]
                label = parts[3] if len(parts) > 3 else None
                try:
                    val = float(val_s)
                except ValueError:
                    continue
                existing = out["fred"].get(key, {})
                out["fred"][key] = {
                    "value":     val,
                    "prev":      existing.get("value"),   # prior FRED value if any
                    "date":      date,
                    "prev_date": existing.get("date"),
                    "label":     label or existing.get("label", key),
                    "kind":      "rate",
                    "series_id": "manual override",
                }
                manual_count += 1
        if manual_count:
            log(f"  Manual overrides applied: {manual_count}")

    # Markets (equity indices, VIX)
    for key, (ticker, label, region) in YF_MARKETS.items():
        data = yf_latest(ticker)
        if data:
            data["label"]  = label
            data["region"] = region
            out["markets"][key] = data
    log(f"  Markets: fetched {len(out['markets'])} / {len(YF_MARKETS)} indices")

    # FX
    for key, (ticker, label) in YF_FX.items():
        data = yf_latest(ticker)
        if data:
            data["label"] = label
            out["fx"][key] = data
    log(f"  FX:      fetched {len(out['fx'])} / {len(YF_FX)} pairs")

    # Live sovereign yields (US via yfinance, UK via BoE IADB)
    for key, (ticker, label, divisor) in LIVE_YIELDS_YF.items():
        d = _yf_yield_latest(ticker, divide_by=divisor)
        if d:
            d["label"] = label
            out["yields_live"][key] = d
    for key, (series_code, label) in LIVE_YIELDS_BOE.items():
        d = _boe_yield_latest(series_code)
        if d:
            d["label"] = label
            out["yields_live"][key] = d
    log(f"  Yields:  fetched {len(out['yields_live'])} live")

    # Finnhub economic calendar — release events with actual/estimate/previous.
    # Next 21 days + recent releases, GB + US + EU + DE + FR.
    out["events"] = _finnhub_events(days_ahead=21)
    log(f"  Events:  fetched {len(out['events'])} upcoming releases")

    # Release overrides: when Finnhub has a freshly-released `actual` that matches
    # one of our tracked FRED series, bump the FRED value to the release value.
    # FRED mirrors typically lag 1-4 weeks; this keeps the Economic Indicators
    # table current without waiting for FRED.
    overrides_applied = _apply_release_overrides(out)
    if overrides_applied:
        log(f"  Release overrides applied: {overrides_applied}")

    OUT_FILE.write_text(json.dumps(out, indent=2), encoding="utf-8")
    log(f"  wrote {OUT_FILE}")


if __name__ == "__main__":
    main()
