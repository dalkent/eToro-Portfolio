#!/usr/bin/env python3
"""
generate_dashboard.py
─────────────────────
Generates a self-contained HTML portfolio dashboard from eToro_Master.xlsx.
Fetches live prices via yfinance and computes P&L, ROI, and valuation signals.

Usage:
    python scripts/generate_dashboard.py

Output:
    eToro_dashboard.html — open directly in any browser. No server required.
"""

import sys
import json
from pathlib import Path
from datetime import datetime

import openpyxl
try:
    from editorial_theme import CSS as _EDITORIAL_CSS, FONTS_LINK as _FONTS_LINK, nav_html as _nav_html, THEME_JS as _THEME_JS
except ImportError:
    from scripts.editorial_theme import CSS as _EDITORIAL_CSS, FONTS_LINK as _FONTS_LINK, nav_html as _nav_html, THEME_JS as _THEME_JS

BASE_DIR  = Path(__file__).parent.parent
DATA_DIR  = BASE_DIR / "data"
MASTER    = DATA_DIR / "eToro_Master.xlsx"
MATCH_CSV = DATA_DIR / "etoro_portfolio_tickermatch.csv"
OUTPUT    = BASE_DIR / "dashboards" / "eToro_dashboard.html"


def load_ticker_buckets() -> dict[str, str]:
    """Load eToro ticker → asset-type bucket mapping from the match CSV.
    Columns: Asset_ID, Ticker_ID, Ticker, Market, Asset (bucket).
    Buckets: UK Equity | International Equity | AI | Corp Bonds | Crypto.
    """
    import csv
    buckets: dict[str, str] = {}
    if not MATCH_CSV.exists():
        return buckets
    with open(MATCH_CSV, encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)  # skip header
        for row in reader:
            if len(row) >= 5:
                ticker = row[2].strip()
                bucket = row[4].strip()
                if ticker and bucket:
                    buckets[ticker] = bucket
    return buckets


# ── Helpers ───────────────────────────────────────────────────────────────────

def signal_for(vr):
    """Returns (label, hex_colour) for a given value ratio."""
    if vr is None:
        return "N/A", "#6b7280"
    if vr >= 1.25:
        return "Strong Buy", "#10b981"
    if vr >= 1.10:
        return "Buy", "#34d399"
    if vr >= 0.90:
        return "Fair Value", "#f59e0b"
    if vr >= 0.75:
        return "Sell", "#f97316"
    return "Strong Sell", "#ef4444"

def fmt_usd(v):
    if v is None:
        return "—"
    return f"${v:,.2f}"

def fmt_pct(v):
    if v is None:
        return "—"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.1f}%"

def fmt_vr(v):
    if v is None:
        return "—"
    return f"{v:.3f}"


# ── Load from JSON cache (produced by sync_xlsx_to_vault.py) ───────────────────

import json as _json

MASTER_JSON = BASE_DIR / "data" / "etoro_master.json"


def _f(v, default=0.0):
    if v in (None, "", "—"):
        return default
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return default


def load_excel():
    """Load all eToro data from the JSON cache written by sync_xlsx_to_vault.py.

    The name is kept for backward compatibility; no Excel file is touched here.
    """
    if not MASTER_JSON.exists():
        print(f"  WARNING: {MASTER_JSON.name} missing. Run sync_xlsx_to_vault.py to build it.")
        return [], [], {}, 0.0, 1.34, []

    print(f"  Reading {MASTER_JSON} ...")
    data = _json.loads(MASTER_JSON.read_text(encoding="utf-8"))
    sheets = data.get("sheets") or {}
    ass    = data.get("assumptions") or {}

    rates       = ass.get("rates") or {}
    gbpusd      = _f(rates.get("GBP/USD"), default=1.34)
    valuations  = ass.get("valuations") or []

    assumptions = {}
    for obj in valuations:
        ticker = str(obj.get("Ticker") or "").strip()
        if not ticker:
            continue
        blended = obj.get("Blended Target (GBP/USD)") or obj.get("Blended Target")
        for k in obj:
            if k and k.startswith("Blended Target"):
                blended = obj.get(k)
                break
        if blended not in (None, "", "—"):
            try:
                assumptions[ticker] = {"blended": float(blended)}
            except (ValueError, TypeError):
                pass

    # Portfolio holdings
    holdings = []
    cash_balance = 0.0
    for obj in (sheets.get("portfolio") or {}).get("objects") or []:
        company = str(obj.get("Company Name") or "").strip()
        if not company:
            continue
        if company.upper() == "CASH":
            cash_balance = _f(obj.get("Invested (USD)"))
            continue
        if "GRAND TOTAL" in company.upper():
            continue
        divs = sum(_f(obj.get(k)) for k in
                   ("Div 2023 (USD)", "Div 2024 (USD)", "Div 2025 (USD)", "Div 2026 (USD)"))
        holdings.append({
            "company":    company,
            "ticker":     str(obj.get("eToro Ticker") or "").strip(),
            "yahoo":      str(obj.get("Yahoo Ticker") or obj.get("eToro Ticker") or "").strip(),
            "sector":     str(obj.get("Sector") or "Other").strip(),
            "currency":   str(obj.get("Currency") or "USD").strip(),
            "units":      _f(obj.get("Units Held")),
            "invested":   _f(obj.get("Invested (USD)")),
            "total_divs": divs,
        })

    # Ticker/company/sector lookup from Tickers sheet
    ticker_names: dict[str, str] = {}
    ticker_sectors: dict[str, str] = {}
    for obj in (sheets.get("tickers") or {}).get("objects") or []:
        company_t = str(obj.get("Company / Name") or "").strip() or None
        sector_t  = str(obj.get("Sector") or "").strip() or None
        for key_col in ("FTSE Ticker (.L)", "eToro Ticker", "Yahoo Finance Ticker"):
            key = str(obj.get(key_col) or "").strip()
            if not key:
                continue
            if company_t:
                ticker_names[key] = company_t
            if sector_t:
                ticker_sectors[key] = sector_t

    # Watchlist
    watchlist = []
    for obj in (sheets.get("watchlist") or {}).get("objects") or []:
        ticker = str(obj.get("eToro Ticker") or "").strip()
        if not ticker:
            continue
        yahoo = str(obj.get("Yahoo Ticker") or ticker).strip()
        company = (
            str(obj.get("Company / Name") or "").strip()
            or ticker_names.get(ticker)
            or ticker_names.get(yahoo)
            or ticker
        )
        sector = (
            str(obj.get("Sector") or "").strip()
            or ticker_sectors.get(ticker)
            or ticker_sectors.get(yahoo)
            or ""
        )
        currency = str(obj.get("Currency") or "GBp").strip()
        watchlist.append({
            "company":  company,
            "ticker":   ticker,
            "yahoo":    yahoo,
            "sector":   sector,
            "currency": currency,
        })

    # Closed positions
    closed = []
    for obj in (sheets.get("closed_positions") or {}).get("objects") or []:
        ticker = str(obj.get("Ticker") or "").strip()
        if not ticker:
            continue
        invested = _f(obj.get("Invested (USD)"))
        sale     = _f(obj.get("Sale Value (USD)"))
        # Sum the year columns directly. The "Total Divs (USD)" column in the
        # spreadsheet is an Excel formula (=SUM of year cols), and the JSON
        # export doesn't preserve formula results — it stores 0. Reading the
        # year cols directly avoids that gap.
        divs     = sum(_f(obj.get(k)) for k in
                       ("Div 2023 (USD)", "Div 2024 (USD)",
                        "Div 2025 (USD)", "Div 2026 (USD)"))
        date_sold = obj.get("Date Sold")
        if isinstance(date_sold, str) and date_sold:
            try:
                from datetime import datetime as _dt
                date_sold = _dt.strptime(date_sold[:10], "%Y-%m-%d").date()
            except ValueError:
                date_sold = None
        closed.append({
            "ticker":    ticker,
            "invested":  invested,
            "sale":      sale,
            "divs":      divs,
            "date_sold": date_sold,
            "pnl":       (sale - invested) + divs,
        })

    return holdings, watchlist, assumptions, cash_balance, gbpusd, closed


# ── Fetch live prices ─────────────────────────────────────────────────────────

def fetch_market_data(all_items, holdings, days_ahead=14):
    """Single yfinance pass: prices, upcoming dividends, earnings, 52wk ranges, dividend income."""
    try:
        import yfinance as yf
    except ImportError:
        print("  Warning: yfinance not installed — run: pip install yfinance")
        return {}, [], [], {}, {}

    # Silence yfinance's chatter about missing fundamentals (expected for ETFs / crypto).
    import logging
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)

    from datetime import date, timedelta
    today = date.today()
    cutoff = today + timedelta(days=days_ahead)

    YF_OVERRIDES = {"BTC": "BTC-USD", "Roku": "ROKU"}
    currency_map = {i.get("yahoo", ""): i.get("currency", "") for i in all_items}
    holding_yahoos = {h["yahoo"] for h in holdings if h.get("yahoo")}
    # Sweep divs+earnings on every ticker we know (portfolio + watchlist).
    divs_earnings_yahoos = {i["yahoo"] for i in all_items if i.get("yahoo")}

    raw_to_yf = {}
    for i in all_items:
        y = i.get("yahoo", "")
        if y:
            raw_to_yf[y] = YF_OVERRIDES.get(y, y)

    # Lookup tables for all items (portfolio + watchlist)
    h_by_yahoo = {}
    for h in holdings:
        y = h.get("yahoo", "")
        if y:
            h_by_yahoo[y] = h
    all_by_yahoo = {}
    for i in all_items:
        y = i.get("yahoo", "")
        if y and y not in all_by_yahoo:
            all_by_yahoo[y] = i

    prices = {}
    daily_changes = {}  # yahoo -> {prev, current, change_pct}
    perf_data = {}      # yahoo -> {ytd_pct, year_pct}
    upcoming_divs = []
    upcoming_earnings = []
    range_data = {}     # yahoo -> {high_52w, low_52w, pct_of_range}
    div_income = {}     # yahoo -> {annual_rate, div_yield}

    total = len(raw_to_yf)
    print(f"  Fetching market data for {total} tickers (single pass) ...")

    for idx, (orig, yf_t) in enumerate(raw_to_yf.items(), 1):
        try:
            t_obj = yf.Ticker(yf_t)

            # ── Price + daily change + YTD + 1yr ──
            hist = t_obj.history(period="1y")
            if not hist.empty:
                price = float(hist["Close"].iloc[-1])
                yf_currency = (t_obj.info.get("currency") or "").upper()
                needs_x100 = currency_map.get(orig) == "GBp" and yf_currency not in ("GBP", "GBX", "GBP", "")
                if needs_x100:
                    price = price * 100
                prices[orig] = price
                if len(hist) >= 2:
                    prev = float(hist["Close"].iloc[-2])
                    if needs_x100:
                        prev = prev * 100
                    chg_pct = ((price - prev) / prev * 100) if prev else 0
                    daily_changes[orig] = {"prev": prev, "current": price, "change_pct": round(chg_pct, 2)}
                # YTD and 1-year (use raw close values; ratio cancels units)
                try:
                    year_start = date(today.year, 1, 1)
                    ytd_series = hist[hist.index.date >= year_start]
                    ytd_start_raw = float(ytd_series["Close"].iloc[0]) if not ytd_series.empty else float(hist["Close"].iloc[-1])
                    year_ago_raw = float(hist["Close"].iloc[0])
                    current_raw = float(hist["Close"].iloc[-1])
                    ytd_pct = (current_raw - ytd_start_raw) / ytd_start_raw * 100 if ytd_start_raw else 0
                    year_pct = (current_raw - year_ago_raw) / year_ago_raw * 100 if year_ago_raw else 0
                    perf_data[orig] = {"ytd_pct": round(ytd_pct, 2), "year_pct": round(year_pct, 2)}
                except Exception:
                    pass

            info = t_obj.info or {}

            # ── 52-week range ──
            high = info.get("fiftyTwoWeekHigh")
            low = info.get("fiftyTwoWeekLow")
            if high and low and high > low:
                raw_price = prices.get(orig)
                if raw_price:
                    pct = (raw_price - low) / (high - low) * 100
                    range_data[orig] = {"high": high, "low": low, "pct_of_range": round(pct, 1)}

            # ── Dividend yield (all items) ──
            annual_rate = info.get("dividendRate")
            div_yield = info.get("dividendYield")
            if annual_rate and annual_rate > 0:
                div_income[orig] = {
                    "annual_rate": annual_rate,
                    "div_yield": round(div_yield, 2) if div_yield else None,
                }

            # ── Portfolio + Watchlist: upcoming dividends, earnings ──
            if orig in divs_earnings_yahoos:
                item = all_by_yahoo.get(orig, {})
                is_watchlist = orig not in holding_yahoos

                # Upcoming dividends
                try:
                    divs = t_obj.dividends
                    if not divs.empty:
                        last_date = divs.index[-1].date()
                        last_amount = float(divs.iloc[-1])
                        if len(divs) >= 2:
                            gaps = [(divs.index[i] - divs.index[i-1]).days
                                    for i in range(1, min(len(divs), 6))]
                            avg_gap = sum(gaps) / len(gaps)
                            next_date = last_date + timedelta(days=int(avg_gap))
                        else:
                            next_date = last_date + timedelta(days=365)
                        try:
                            cal = t_obj.calendar or {}
                            cal_ex = cal.get("Ex-Dividend Date")
                            if cal_ex and cal_ex >= today:
                                next_date = cal_ex
                            # Earnings
                            for ed in cal.get("Earnings Date", []):
                                if today <= ed <= cutoff:
                                    upcoming_earnings.append({
                                        "ticker": item.get("ticker", orig),
                                        "company": item.get("company", orig),
                                        "date": ed,
                                        "value": item.get("current_value") or item.get("invested", 0),
                                        "watchlist": is_watchlist,
                                    })
                                    break
                        except Exception:
                            pass
                        if today <= next_date <= cutoff:
                            upcoming_divs.append({
                                "ticker": item.get("ticker", orig),
                                "company": item.get("company", orig),
                                "currency": item.get("currency", "USD"),
                                "ex_date": next_date,
                                "amount": last_amount,
                                "units": item.get("units", 0),
                                "div_yield": div_income.get(orig, {}).get("div_yield"),
                                "watchlist": is_watchlist,
                            })
                    else:
                        # No dividend history — still check earnings
                        try:
                            cal = t_obj.calendar or {}
                            for ed in cal.get("Earnings Date", []):
                                if today <= ed <= cutoff:
                                    upcoming_earnings.append({
                                        "ticker": item.get("ticker", orig),
                                        "company": item.get("company", orig),
                                        "date": ed,
                                        "value": item.get("current_value") or item.get("invested", 0),
                                        "watchlist": is_watchlist,
                                    })
                                    break
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

        if idx % 25 == 0:
            print(f"    {idx}/{total}")

    upcoming_divs.sort(key=lambda x: x["ex_date"])
    upcoming_earnings.sort(key=lambda x: x["date"])
    # ── Benchmarks: FTSE 100 & S&P 500 ──
    benchmarks = {}
    for name, symbol in [("FTSE 100", "^FTSE"), ("S&P 500", "^GSPC")]:
        try:
            bt = yf.Ticker(symbol)
            hist = bt.history(period="1y")
            if not hist.empty:
                current = float(hist["Close"].iloc[-1])
                year_ago = float(hist["Close"].iloc[0])
                # YTD: find first trading day of this calendar year
                from datetime import date as _date
                year_start = _date(today.year, 1, 1)
                ytd_series = hist[hist.index.date >= year_start]
                ytd_start = float(ytd_series["Close"].iloc[0]) if not ytd_series.empty else current
                benchmarks[name] = {
                    "ytd_pct": round((current - ytd_start) / ytd_start * 100, 2),
                    "year_pct": round((current - year_ago) / year_ago * 100, 2),
                }
        except Exception:
            pass

    print(f"  Prices: {len(prices)} | Divs: {len(upcoming_divs)} | Earnings: {len(upcoming_earnings)} | 52wk: {len(range_data)} | Div income: {len(div_income)} | Daily chg: {len(daily_changes)} | Perf: {len(perf_data)} | Benchmarks: {len(benchmarks)}")
    return prices, upcoming_divs, upcoming_earnings, range_data, div_income, daily_changes, benchmarks, perf_data


# ── Compute derived metrics ───────────────────────────────────────────────────

def enrich(holdings, watchlist, assumptions, cash, gbpusd, prices, closed=None):
    """Attach live_price, current_value, pnl, roi, target, vr, signal to each holding."""
    closed = closed or []

    def _price_usd(item, raw_price):
        if raw_price is None:
            return None
        if item["currency"] == "GBp":
            return (raw_price / 100) * gbpusd   # pence → GBP → USD
        return raw_price                          # already USD

    def _target_and_vr(ticker, currency, raw_price):
        assum = assumptions.get(ticker, {})
        target = assum.get("blended")
        if target is None:
            return None, None

        if raw_price is None:
            return target, None
        # Compare in same unit: GBP for .L stocks, USD otherwise
        if currency == "GBp":
            current = raw_price / 100
        else:
            current = raw_price
        vr = target / current if current else None
        return target, vr

    total_invested = sum(h["invested"] for h in holdings) + cash
    total_value    = cash
    total_divs     = 0.0

    for h in holdings:
        raw = prices.get(h["yahoo"])
        h["live_price"] = raw
        pusd = _price_usd(h, raw)
        h["current_value"] = (h["units"] * pusd) if pusd else h["invested"]
        h["pnl"] = h["current_value"] - h["invested"]
        h["roi"] = (h["pnl"] / h["invested"] * 100) if h["invested"] else 0

        total_value += h["current_value"]
        total_divs  += h["total_divs"]

        target, vr = _target_and_vr(h["ticker"], h["currency"], raw)
        h["target"] = target
        h["vr"]     = vr
        h["signal"], h["signal_color"] = signal_for(vr)

    for w in watchlist:
        raw = prices.get(w["yahoo"])
        w["live_price"] = raw
        target, vr = _target_and_vr(w["ticker"], w["currency"], raw)
        w["target"] = target
        w["vr"]     = vr
        w["signal"], w["signal_color"] = signal_for(vr)
        if raw:
            if w["currency"] == "GBp":
                w["live_price_display"] = f"£{raw / 100:.2f}"
            else:
                w["live_price_display"] = f"${raw:.2f}"
        else:
            w["live_price_display"] = "—"

    # Closed positions contribution
    closed_invested = sum(c["invested"] for c in closed)
    closed_sale     = sum(c["sale"] for c in closed)
    closed_divs     = sum(c["divs"] for c in closed)
    realized_pnl    = (closed_sale - closed_invested) + closed_divs

    capital_pnl  = (total_value - total_invested) + (closed_sale - closed_invested)
    total_return = capital_pnl + total_divs + closed_divs
    total_base   = total_invested + closed_invested
    total_roi    = (total_return / total_base * 100) if total_base else 0

    # Asset type classification
    BUFFER_TICKERS      = {"LQDE.L", "IGLT.L", "SLXX.L"}
    BUFFER_SECTORS      = {"Corp Bonds", "Government Bonds", "Fixed Income"}
    INTL_EQUITY_TICKERS = {"CCI", "CVS", "UMC", "PEP", "ONON", "VZ"}  # explicit whitelist

    def classify(h):
        t = h["ticker"]
        s = h["sector"] or ""
        if t in BUFFER_TICKERS or s in BUFFER_SECTORS:
            return "Buffer (Bonds & Cash)"
        if h["currency"] == "GBp" or t.endswith(".L"):
            return "UK Equities"
        if t in INTL_EQUITY_TICKERS:
            return "US & International Equity"
        # everything else non-UK is Crypto & Growth
        return "Crypto & Growth"

    asset_types = {}
    for h in holdings:
        at = classify(h)
        h["asset_type"] = at
        asset_types[at] = asset_types.get(at, 0.0) + h["current_value"]

    # Add cash to buffer
    asset_types["Buffer (Bonds & Cash)"] = asset_types.get("Buffer (Bonds & Cash)", 0.0) + cash

    # Target allocations (midpoints of ranges)
    targets = {
        "UK Equities":               70.0,   # 65-75% midpoint
        "Buffer (Bonds & Cash)":     12.5,   # 10-15% midpoint
        "Crypto & Growth":           12.5,   # 10-15% midpoint
        "US & International Equity":  5.0,   # 2.5-7.5% midpoint
    }

    holdings.sort(key=lambda x: x["current_value"], reverse=True)

    summary = {
        "total_invested":  total_invested + closed_invested,
        "open_invested":   total_invested,
        "total_value":     total_value,
        "cash":            cash,
        "capital_pnl":     capital_pnl,
        "total_divs":      total_divs + closed_divs,
        "total_return":    total_return,
        "total_roi":       total_roi,
        "realized_pnl":    realized_pnl,
        "closed_count":    len(closed),
        "closed_invested": closed_invested,
        "closed_sale":     closed_sale,
        "asset_types":     asset_types,
        "targets":         targets,
        "generated_at":    datetime.now().strftime("%d %b %Y %H:%M"),
        "prices_live":     bool(prices),
    }

    return summary


# ── HTML generation ───────────────────────────────────────────────────────────

ASSET_COLOURS = {
    "UK Equities":               "#3b82f6",
    "Buffer (Bonds & Cash)":     "#10b981",
    "Crypto & Growth":           "#f59e0b",
    "US & International Equity": "#8b5cf6",
}

TECH_AI_TICKERS = {
    "AMD","SHOP","U","TSLA","RKLB","ACHR","RBLX","TEM","KTOS","IRDM","NVDA","META","GOOGL","MSFT","AMZN","BMNR"
}

def build_html(holdings, watchlist, summary, upcoming_divs=None, upcoming_earnings=None, range_data=None, div_income=None, daily_changes=None, benchmarks=None, perf_data=None, closed=None):

    # Asset type table data
    total_val = summary["total_value"] or 1
    all_labels = list(summary["targets"].keys())
    actual_pcts  = [round(summary["asset_types"].get(l, 0) / total_val * 100, 1) for l in all_labels]
    actual_vals  = [round(summary["asset_types"].get(l, 0), 2) for l in all_labels]
    target_pcts  = [summary["targets"][l] for l in all_labels]
    chart_colours = [ASSET_COLOURS.get(l, "#6366f1") for l in all_labels]

    def alloc_rows():
        rows = []
        for i, label in enumerate(all_labels):
            act  = actual_pcts[i]
            tgt  = target_pcts[i]
            diff = round(act - tgt, 1)
            val  = actual_vals[i]
            colour = ASSET_COLOURS.get(label, "#6366f1")
            diff_str  = f"+{diff}%" if diff > 0 else f"{diff}%"
            diff_cls  = "pos" if diff >= -2 else "neg"
            status    = "On target" if abs(diff) <= 2 else ("Overweight" if diff > 0 else "Underweight")
            rows.append(f"""
            <tr>
              <td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{colour};margin-right:7px;"></span>{label}</td>
              <td class="num">{fmt_usd(val)}</td>
              <td class="num">{act}%</td>
              <td class="num">{tgt}%</td>
              <td class="num {diff_cls}">{diff_str}</td>
              <td class="num">{status}</td>
            </tr>""")
        return "\n".join(rows)

    # Signals: UK strong sells only (portfolio)
    uk_strong_sells = sorted([
        h for h in holdings
        if h.get("signal") == "Strong Sell"
        and (h["ticker"].endswith(".L") or h["currency"] == "GBp")
    ], key=lambda x: x.get("vr") or 99)

    # Signals: Tech & AI positions needing action
    tech_add  = [h for h in holdings if h["ticker"] in TECH_AI_TICKERS and h["current_value"] < 15]
    tech_sell = [h for h in holdings if h["ticker"] in TECH_AI_TICKERS and h["current_value"] > 35]

    # Signals: watchlist strong buys only, sorted by VR desc
    watchlist_strong_buys = sorted([
        w for w in watchlist
        if w.get("signal") == "Strong Buy"
    ], key=lambda w: w.get("vr") or 0, reverse=True)

    def strong_buy_rows():
        if not watchlist_strong_buys:
            return '<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:8px">None</td></tr>'
        rows = []
        for w in watchlist_strong_buys:
            tgt = (f'£{w["target"]:.2f}' if w["currency"] == "GBp" else f'${w["target"]:.2f}') if w["target"] else "—"
            rows.append(f"""
            <tr>
              <td><strong>{w["ticker"]}</strong><br><small>{w["company"][:28]}</small></td>
              <td><small>{w["sector"] or "—"}</small></td>
              <td class="num">{w.get("live_price_display","—")}</td>
              <td class="num">{tgt}</td>
              <td class="num pos">{fmt_vr(w["vr"])}</td>
            </tr>""")
        return "\n".join(rows)

    def strong_sell_rows():
        if not uk_strong_sells:
            return '<tr><td colspan="5" style="color:var(--muted);text-align:center;padding:8px">None</td></tr>'
        rows = []
        for h in uk_strong_sells:
            price_str = f'£{h["live_price"]/100:.2f}' if h["currency"] == "GBp" else fmt_usd(h["live_price"])
            rows.append(f"""
            <tr>
              <td><strong>{h["ticker"]}</strong><br><small>{h["company"][:28]}</small></td>
              <td><small>{h["sector"] or "—"}</small></td>
              <td class="num">{price_str}</td>
              <td class="num">{fmt_usd(h["current_value"])}</td>
              <td class="num neg">{fmt_vr(h["vr"])}</td>
            </tr>""")
        return "\n".join(rows)

    _wl_pill = (
        '<span style="font-family:var(--f-mono);font-size:8px;letter-spacing:.08em;'
        'text-transform:uppercase;background:var(--bg-sunken);color:var(--ink-3);'
        'padding:1px 5px;border-radius:3px;margin-left:6px;vertical-align:middle;'
        '" title="Watchlist">WL</span>'
    )

    def dividend_rows():
        divs = upcoming_divs or []
        if not divs:
            return '<tr><td colspan="6" style="color:var(--muted);text-align:center;padding:8px">No ex-dividends in the next 14 days</td></tr>'
        rows = []
        for d in divs:
            ex = d["ex_date"].strftime("%a %d %b")
            if d["currency"] == "GBp":
                amt_str = f'{d["amount"]:.2f}p'
                est = d["amount"] * d["units"] / 100
                est_str = f'£{est:.2f}' if d["units"] else "—"
            else:
                amt_str = f'${d["amount"]:.4f}'
                est = d["amount"] * d["units"]
                est_str = f'${est:.2f}' if d["units"] else "—"
            yld = f'{d["div_yield"]:.1f}%' if d["div_yield"] else "—"
            wl = _wl_pill if d.get("watchlist") else ""
            rows.append(f"""
            <tr>
              <td><strong>{d["ticker"]}</strong>{wl}<br><small>{d["company"][:28]}</small></td>
              <td class="num">{ex}</td>
              <td class="num">{amt_str}</td>
              <td class="num">{yld}</td>
              <td class="num">{est_str}</td>
            </tr>""")
        return "\n".join(rows)

    def earnings_rows():
        items = upcoming_earnings or []
        if not items:
            return '<tr><td colspan="3" style="color:var(--muted);text-align:center;padding:8px">No earnings in the next 14 days</td></tr>'
        # Build live ticker→current_value lookup from enriched holdings.
        # NOTE: upcoming_earnings is built in fetch_market_data() before enrich() runs,
        # so e["value"] is the *invested* amount, not current value. Override here.
        live_value = {h["ticker"]: h.get("current_value") for h in holdings}
        rows = []
        for e in items:
            dt = e["date"].strftime("%a %d %b")
            val = live_value.get(e["ticker"]) if not e.get("watchlist") else None
            val_str = fmt_usd(val) if val else "—"
            wl = _wl_pill if e.get("watchlist") else ""
            rows.append(f"""
            <tr>
              <td><strong>{e["ticker"]}</strong>{wl}<br><small>{e["company"][:28]}</small></td>
              <td class="num">{dt}</td>
              <td class="num">{val_str}</td>
            </tr>""")
        return "\n".join(rows)

    def range_alert_rows():
        rd = range_data or {}
        near_high = []
        near_low = []
        for h in holdings:
            y = h["yahoo"]
            r = rd.get(y)
            if not r:
                continue
            pct = r["pct_of_range"]
            entry = {**h, "pct": pct, "high": r["high"], "low": r["low"]}
            if pct >= 90:
                near_high.append(entry)
            elif pct <= 10:
                near_low.append(entry)
        near_high.sort(key=lambda x: x["pct"], reverse=True)
        near_low.sort(key=lambda x: x["pct"])
        if not near_high and not near_low:
            return '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:8px">No stocks near 52-week extremes</td></tr>'
        rows = []
        for e in near_high:
            tag = '<span class="badge" style="background:#10b981">Near High</span>'
            rows.append(f"""
            <tr>
              <td><strong>{e["ticker"]}</strong><br><small>{e["company"][:28]}</small></td>
              <td class="num">{e["pct"]:.0f}%</td>
              <td>{tag}</td>
            </tr>""")
        for e in near_low:
            tag = '<span class="badge" style="background:#ef4444">Near Low</span>'
            rows.append(f"""
            <tr>
              <td><strong>{e["ticker"]}</strong><br><small>{e["company"][:28]}</small></td>
              <td class="num">{e["pct"]:.0f}%</td>
              <td>{tag}</td>
            </tr>""")
        return "\n".join(rows)

    def benchmark_rows():
        def _cls(v):
            if v is None:
                return ""
            return "pos" if v >= 0 else "neg"
        rows = []
        # Portfolio row first (weighted look-through)
        if portfolio_ytd is not None or portfolio_1y is not None:
            rows.append(f"""
            <tr style="font-weight:600;background:rgba(59,130,246,0.08)">
              <td><strong>Portfolio</strong> <small>(weighted)</small></td>
              <td class="num {_cls(portfolio_ytd)}">{fmt_pct(portfolio_ytd)}</td>
              <td class="num {_cls(portfolio_1y)}">{fmt_pct(portfolio_1y)}</td>
            </tr>""")
        for name, d in bm.items():
            ytd = d.get("ytd_pct")
            yr = d.get("year_pct")
            rows.append(f"""
            <tr>
              <td>{name}</td>
              <td class="num {_cls(ytd)}">{fmt_pct(ytd)}</td>
              <td class="num {_cls(yr)}">{fmt_pct(yr)}</td>
            </tr>""")
        if not rows:
            return '<tr><td colspan="3" style="color:var(--muted);text-align:center;padding:8px">No benchmark data</td></tr>'
        return "\n".join(rows)

    def currency_rows():
        if not ccy_items:
            return '<tr><td colspan="3" style="color:var(--muted);text-align:center;padding:8px">No data</td></tr>'
        rows = []
        for ccy, val in ccy_items:
            pct = val / total_val * 100
            rows.append(f"""
            <tr>
              <td>{ccy}</td>
              <td class="num">{fmt_usd(val)}</td>
              <td class="num">{pct:.1f}%</td>
            </tr>""")
        return "\n".join(rows)

    def top_movers_rows():
        dc = daily_changes or {}
        movers = []
        for h in holdings:
            d = dc.get(h["yahoo"])
            if not d:
                continue
            movers.append({**h, "change_pct": d["change_pct"]})
        movers.sort(key=lambda x: abs(x["change_pct"]), reverse=True)
        top = movers[:10]
        if not top:
            return '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:8px">No price data</td></tr>'
        rows = []
        for m in top:
            cls = "pos" if m["change_pct"] >= 0 else "neg"
            sign = "+" if m["change_pct"] >= 0 else ""
            impact = m["current_value"] * m["change_pct"] / (100 + m["change_pct"]) if m["change_pct"] != -100 else 0
            rows.append(f"""
            <tr>
              <td><strong>{m["ticker"]}</strong><br><small>{m["company"][:28]}</small></td>
              <td class="num {cls}">{sign}{m["change_pct"]:.1f}%</td>
              <td class="num {cls}">{fmt_usd(impact)}</td>
            </tr>""")
        return "\n".join(rows)

    def sector_pnl_rows():
        sectors = {}
        for h in holdings:
            s = h.get("sector") or "Other"
            if s not in sectors:
                sectors[s] = {"invested": 0, "value": 0, "pnl": 0}
            sectors[s]["invested"] += h["invested"]
            sectors[s]["value"] += h["current_value"]
            sectors[s]["pnl"] += h["pnl"]
        items = sorted(sectors.items(), key=lambda x: x[1]["pnl"], reverse=True)
        rows = []
        for s, d in items:
            roi = (d["pnl"] / d["invested"] * 100) if d["invested"] else 0
            cls = "pos" if d["pnl"] >= 0 else "neg"
            rows.append(f"""
            <tr>
              <td>{s}</td>
              <td class="num">{fmt_usd(d["value"])}</td>
              <td class="num {cls}">{fmt_usd(d["pnl"])}</td>
              <td class="num {cls}">{fmt_pct(roi)}</td>
            </tr>""")
        return "\n".join(rows)

    def asset_bucket_pnl_rows():
        """Performance grouped by Asset Type bucket (UK Equity / AI / Crypto / International Equity / ...)."""
        bucket_map = load_ticker_buckets()
        # Canonical display order — user explicitly requested these four; show others (e.g. Corp Bonds) after.
        PREFERRED = ["UK Equity", "AI", "Crypto", "International Equity"]

        stats = {}
        for h in holdings:
            b = bucket_map.get(h["ticker"]) or "Other"
            if b not in stats:
                stats[b] = {"count": 0, "invested": 0.0, "value": 0.0,
                            "pnl": 0.0, "divs": 0.0}
            stats[b]["count"]    += 1
            stats[b]["invested"] += h["invested"]
            stats[b]["value"]    += h["current_value"]
            stats[b]["pnl"]      += h["pnl"]
            stats[b]["divs"]     += h["total_divs"]

        # Order: preferred buckets first (in requested order), then any extras alphabetically
        ordered = [b for b in PREFERRED if b in stats]
        ordered += sorted(b for b in stats if b not in PREFERRED)

        rows = []
        totals = {"count": 0, "invested": 0.0, "value": 0.0, "pnl": 0.0, "divs": 0.0}
        for b in ordered:
            d = stats[b]
            total_return = d["pnl"] + d["divs"]
            cap_pct = (d["pnl"] / d["invested"] * 100) if d["invested"] else 0
            roi_pct = (total_return / d["invested"] * 100) if d["invested"] else 0
            cap_cls = "pos" if d["pnl"] >= 0 else "neg"
            roi_cls = "pos" if total_return >= 0 else "neg"
            rows.append(f"""
            <tr>
              <td><strong>{b}</strong></td>
              <td class="num">{d["count"]}</td>
              <td class="num">{fmt_usd(d["invested"])}</td>
              <td class="num">{fmt_usd(d["value"])}</td>
              <td class="num {cap_cls}">{fmt_usd(d["pnl"])}</td>
              <td class="num {cap_cls}">{fmt_pct(cap_pct)}</td>
              <td class="num">{fmt_usd(d["divs"])}</td>
              <td class="num {roi_cls}">{fmt_usd(total_return)}</td>
              <td class="num {roi_cls}">{fmt_pct(roi_pct)}</td>
            </tr>""")
            for k in totals:
                totals[k] += d[k]

        # Totals row
        t_return = totals["pnl"] + totals["divs"]
        t_cap_pct = (totals["pnl"] / totals["invested"] * 100) if totals["invested"] else 0
        t_roi_pct = (t_return / totals["invested"] * 100) if totals["invested"] else 0
        t_cap_cls = "pos" if totals["pnl"] >= 0 else "neg"
        t_roi_cls = "pos" if t_return >= 0 else "neg"
        rows.append(f"""
            <tr style="border-top:2px solid var(--border);font-weight:600">
              <td>Total</td>
              <td class="num">{totals["count"]}</td>
              <td class="num">{fmt_usd(totals["invested"])}</td>
              <td class="num">{fmt_usd(totals["value"])}</td>
              <td class="num {t_cap_cls}">{fmt_usd(totals["pnl"])}</td>
              <td class="num {t_cap_cls}">{fmt_pct(t_cap_pct)}</td>
              <td class="num">{fmt_usd(totals["divs"])}</td>
              <td class="num {t_roi_cls}">{fmt_usd(t_return)}</td>
              <td class="num {t_roi_cls}">{fmt_pct(t_roi_pct)}</td>
            </tr>""")

        return "\n".join(rows)

    def concentration_rows():
        total_val = summary["total_value"] or 1
        flagged = []
        for h in holdings:
            weight = h["current_value"] / total_val * 100
            is_uk = h["currency"] == "GBp" or h["ticker"].endswith(".L")
            threshold = 5.0 if is_uk else 10.0
            if weight >= threshold:
                flagged.append({**h, "weight": weight, "threshold": threshold, "is_uk": is_uk})
        flagged.sort(key=lambda x: x["weight"], reverse=True)
        if not flagged:
            return '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:8px">No concentration flags</td></tr>'
        rows = []
        for f in flagged:
            label = f"UK >{f['threshold']:.0f}%" if f["is_uk"] else f">{f['threshold']:.0f}%"
            rows.append(f"""
            <tr>
              <td><strong>{f["ticker"]}</strong><br><small>{f["company"][:28]}</small></td>
              <td class="num">{f["weight"]:.1f}%</td>
              <td class="num">{fmt_usd(f["current_value"])}</td>
              <td><span class="badge" style="background:#f59e0b">{label}</span></td>
            </tr>""")
        return "\n".join(rows)

    def rebalance_rows():
        total_val = summary["total_value"] or 1
        all_labels = list(summary["targets"].keys())
        actions = []
        for label in all_labels:
            actual_val = summary["asset_types"].get(label, 0)
            target_pct = summary["targets"][label]
            target_val = total_val * target_pct / 100
            diff_val = actual_val - target_val
            actual_pct = actual_val / total_val * 100
            diff_pct = actual_pct - target_pct
            if abs(diff_pct) > 2:
                action = "Reduce" if diff_val > 0 else "Add"
                colour = ASSET_COLOURS.get(label, "#6366f1")
                cls = "neg" if diff_val > 0 else "pos"
                actions.append(f"""
            <tr>
              <td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{colour};margin-right:7px"></span>{label}</td>
              <td class="num {cls}">{fmt_usd(abs(diff_val))}</td>
              <td><span class="badge" style="background:{'#ef4444' if action=='Reduce' else '#10b981'}">{action}</span></td>
            </tr>""")
        if not actions:
            return '<tr><td colspan="3" style="color:var(--muted);text-align:center;padding:8px">Portfolio is balanced</td></tr>'
        return "\n".join(actions)

    def tech_action_rows():
        if not tech_sell and not tech_add:
            return '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:8px">No actions needed</td></tr>'
        rows = []
        for h in tech_sell:
            rows.append(f"""
            <tr>
              <td><strong>{h["ticker"]}</strong><br><small>{h["company"][:28]}</small></td>
              <td class="num">{fmt_usd(h["current_value"])}</td>
              <td class="num {('pos' if h['roi']>=0 else 'neg')}">{fmt_pct(h["roi"])}</td>
              <td><span class="badge" style="background:#ef4444">Sell &gt;$35</span></td>
            </tr>""")
        for h in tech_add:
            rows.append(f"""
            <tr>
              <td><strong>{h["ticker"]}</strong><br><small>{h["company"][:28]}</small></td>
              <td class="num">{fmt_usd(h["current_value"])}</td>
              <td class="num {('pos' if h['roi']>=0 else 'neg')}">{fmt_pct(h["roi"])}</td>
              <td><span class="badge" style="background:#f59e0b">Add &lt;$15</span></td>
            </tr>""")
        return "\n".join(rows)

    def row_class(roi):
        if roi is None:
            return ""
        if roi >= 0:
            return "positive"
        return "negative"

    def holding_rows():
        di = div_income or {}
        rows = []
        for h in holdings:
            signal_badge = (
                f'<span class="badge" style="background:{h["signal_color"]}">'
                f'{h["signal"]}</span>'
            )
            vr_str = fmt_vr(h["vr"])
            pnl_str = fmt_usd(h["pnl"])
            roi_str = fmt_pct(h["roi"])
            val_str = fmt_usd(h["current_value"])
            inv_str = fmt_usd(h["invested"])
            div_str = fmt_usd(h["total_divs"]) if h["total_divs"] else "—"

            d = di.get(h["yahoo"], {})
            yld = d.get("div_yield")
            yld_str = f'{yld:.1f}%' if yld else "—"

            if h["live_price"] is None:
                price_str = '<span class="stale">No price</span>'
            elif h["currency"] == "GBp":
                price_str = f'{h["live_price"]:.2f}p'
            else:
                price_str = f'${h["live_price"]:.2f}'

            pnl_cls = "pos" if h["pnl"] >= 0 else "neg"
            rows.append(f"""
            <tr>
              <td><strong>{h["ticker"]}</strong><br><small>{h["company"][:28]}</small></td>
              <td><small>{h["sector"]}</small></td>
              <td class="num">{inv_str}</td>
              <td class="num">{val_str}</td>
              <td class="num {pnl_cls}">{pnl_str}</td>
              <td class="num {pnl_cls}">{roi_str}</td>
              <td class="num">{div_str}</td>
              <td class="num">{yld_str}</td>
              <td class="num">{price_str}</td>
              <td class="num">{vr_str}</td>
              <td>{signal_badge}</td>
            </tr>""")
        return "\n".join(rows)

    def watchlist_rows():
        di = div_income or {}
        rows = []
        for w in watchlist:
            signal_badge = (
                f'<span class="badge" style="background:{w["signal_color"]}">'
                f'{w["signal"]}</span>'
            )
            if w["target"] is None:
                target_str = "—"
            elif w["currency"] == "GBp":
                target_str = f"£{w['target']:.2f}"
            else:
                target_str = f"${w['target']:.2f}"
            d = di.get(w["yahoo"], {})
            yld = d.get("div_yield")
            yld_str = f'{yld:.1f}%' if yld else "—"
            rows.append(f"""
            <tr>
              <td><strong>{w["ticker"]}</strong><br><small>{w["company"][:30]}</small></td>
              <td><small>{w["sector"] or "—"}</small></td>
              <td class="num">{w.get("live_price_display","—")}</td>
              <td class="num">{target_str}</td>
              <td class="num">{yld_str}</td>
              <td class="num">{fmt_vr(w["vr"])}</td>
              <td>{signal_badge}</td>
            </tr>""")
        return "\n".join(rows)

    pnl_colour  = "#10b981" if summary["capital_pnl"] >= 0 else "#ef4444"
    roi_colour  = "#10b981" if summary["total_roi"]   >= 0 else "#ef4444"
    live_note   = "Live prices via Yahoo Finance" if summary["prices_live"] else "⚠️ Prices unavailable — install yfinance"

    # Currency exposure (USD terms)
    ccy_totals = {}
    for h in holdings:
        ccy = "GBP" if h["currency"] == "GBp" else h["currency"]
        ccy_totals[ccy] = ccy_totals.get(ccy, 0) + h["current_value"]
    ccy_totals["USD (Cash)"] = ccy_totals.get("USD", 0)  # separate cash? simpler: just add
    ccy_totals.pop("USD (Cash)", None)
    ccy_totals["USD"] = ccy_totals.get("USD", 0) + summary["cash"]
    total_val = summary["total_value"] or 1
    ccy_items = sorted(ccy_totals.items(), key=lambda x: x[1], reverse=True)

    # Benchmark comparison
    bm = benchmarks or {}

    # Weighted portfolio YTD and 1-year (look-through + closed positions)
    from datetime import date as _date_lt, timedelta as _td_lt
    pd_data = perf_data or {}
    today_lt = _date_lt.today()
    year_start_lt = _date_lt(today_lt.year, 1, 1)
    year_ago_lt = today_lt - _td_lt(days=365)

    equity_value = sum(h["current_value"] for h in holdings if pd_data.get(h["yahoo"]))
    # Open-holdings weighted dollar return
    open_ytd_dollars = sum(h["current_value"] * pd_data[h["yahoo"]]["ytd_pct"] / 100 for h in holdings if pd_data.get(h["yahoo"]))
    open_1y_dollars  = sum(h["current_value"] * pd_data[h["yahoo"]]["year_pct"] / 100 for h in holdings if pd_data.get(h["yahoo"]))

    # Convert current_value back to approximate start-of-period value (avoids double-counting the return)
    # Start value = current / (1 + return/100)
    open_ytd_base = sum(h["current_value"] / (1 + pd_data[h["yahoo"]]["ytd_pct"]/100) for h in holdings if pd_data.get(h["yahoo"]) and pd_data[h["yahoo"]]["ytd_pct"] != -100)
    open_1y_base  = sum(h["current_value"] / (1 + pd_data[h["yahoo"]]["year_pct"]/100) for h in holdings if pd_data.get(h["yahoo"]) and pd_data[h["yahoo"]]["year_pct"] != -100)

    # Closed-position contribution within YTD / 1-year windows
    closed_list = closed or []
    closed_ytd_pnl, closed_ytd_base = 0.0, 0.0
    closed_1y_pnl, closed_1y_base = 0.0, 0.0
    for c in closed_list:
        ds = c.get("date_sold")
        if not ds:
            continue
        # Normalise to date object
        if hasattr(ds, "date"):
            ds = ds.date()
        if ds >= year_start_lt:
            closed_ytd_pnl  += c["pnl"]
            closed_ytd_base += c["invested"]
        if ds >= year_ago_lt:
            closed_1y_pnl  += c["pnl"]
            closed_1y_base += c["invested"]

    portfolio_ytd = None
    portfolio_1y = None
    if open_ytd_base > 0:
        portfolio_ytd = round((open_ytd_dollars + closed_ytd_pnl) / (open_ytd_base + closed_ytd_base) * 100, 2)
    if open_1y_base > 0:
        portfolio_1y = round((open_1y_dollars + closed_1y_pnl) / (open_1y_base + closed_1y_base) * 100, 2)

    # Projected annual income
    di = div_income or {}
    proj_usd, proj_gbp = 0.0, 0.0
    for h in holdings:
        d = di.get(h["yahoo"], {})
        rate = d.get("annual_rate", 0)
        if rate and rate > 0:
            if h["currency"] == "GBp":
                proj_gbp += rate * h.get("units", 0) / 100
            else:
                proj_usd += rate * h.get("units", 0)
    proj_parts = []
    if proj_usd > 0:
        proj_parts.append(f"${proj_usd:,.2f}")
    if proj_gbp > 0:
        proj_parts.append(f"\u00a3{proj_gbp:,.2f}")
    proj_income_str = " + ".join(proj_parts) if proj_parts else "$0"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>eToro Portfolio Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
{_FONTS_LINK}
<style>{_EDITORIAL_CSS}</style>
</head>
<body>
{_nav_html(active="etoro", privacy=True)}

<div class="header">
  <h1>eToro Portfolio Dashboard</h1>
  <div class="header-right">
    Generated: <span id="gen-ts" data-ts="{int(datetime.now().timestamp())}">{summary["generated_at"]}</span><br>
    {live_note}
  </div>
</div>

<!-- In-page section nav -->
<div class="nav-bar">
  <a href="#overview">Overview</a>
  <a href="#signals">Signals</a>
  <a href="#analytics">Analytics</a>
  <a href="#news">News</a>
  <a href="#holdings">Holdings</a>
  <a href="#watchlist">Watchlist</a>
</div>

<div id="stale-banner"></div>

<!-- KPI Summary -->
<div class="kpi-grid" id="overview">
  <div class="kpi">
    <div class="kpi-label">Portfolio Value</div>
    <div class="kpi-value">{fmt_usd(summary["total_value"])}</div>
    <div class="kpi-sub">vs. invested</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Cash Position</div>
    <div class="kpi-value">{fmt_usd(summary["cash"])}</div>
    <div class="kpi-sub">{summary["cash"]/summary["total_value"]*100:.1f}% of portfolio</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Capital P&amp;L</div>
    <div class="kpi-value" style="color:{pnl_colour}">{fmt_usd(summary["capital_pnl"])}</div>
    <div class="kpi-sub">open + realized</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Realized P&amp;L</div>
    <div class="kpi-value" style="color:{'var(--pos)' if summary.get('realized_pnl', 0) >= 0 else 'var(--neg)'}">{fmt_usd(summary.get("realized_pnl", 0))}</div>
    <div class="kpi-sub">{summary.get("closed_count", 0)} closed positions</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Dividends Received</div>
    <div class="kpi-value" style="color:var(--pos)">{fmt_usd(summary["total_divs"])}</div>
    <div class="kpi-sub">all years</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Total Return</div>
    <div class="kpi-value" style="color:{roi_colour}">{fmt_pct(summary["total_roi"])}</div>
    <div class="kpi-sub">Capital + Divs</div>
  </div>
</div>

<!-- Benchmarks & Currency Exposure -->
<div class="signal-grid" style="margin-bottom:16px">
<div class="card">
  <h2>Performance vs Benchmarks</h2>
  <table>
    <thead><tr><th>Benchmark</th><th class="num">YTD</th><th class="num">1 Year</th></tr></thead>
    <tbody>{benchmark_rows()}</tbody>
  </table>
</div>
<div class="card">
  <h2>Currency Exposure</h2>
  <table>
    <thead><tr><th>Currency</th><th class="num">Value (USD)</th><th class="num">% of Portfolio</th></tr></thead>
    <tbody>{currency_rows()}</tbody>
  </table>
</div>
</div>

<!-- Action Signals: 3 tables side by side -->
<div class="signal-grid" id="signals">

  <div class="card">
    <h2>Watchlist — Strong Buys</h2>
    <table>
      <thead>
        <tr>
          <th>Stock</th>
          <th>Sector</th>
          <th class="num">Price</th>
          <th class="num">Target</th>
          <th class="num">VR</th>
        </tr>
      </thead>
      <tbody>
        {strong_buy_rows()}
      </tbody>
    </table>
  </div>

  <div class="card">
    <h2>Portfolio — Strong Sells (UK)</h2>
    <table>
      <thead>
        <tr>
          <th>Stock</th>
          <th>Sector</th>
          <th class="num">Price</th>
          <th class="num">Value</th>
          <th class="num">VR</th>
        </tr>
      </thead>
      <tbody>
        {strong_sell_rows()}
      </tbody>
    </table>
  </div>

  <div class="card">
    <h2>Tech &amp; AI — Actions</h2>
    <table>
      <thead>
        <tr>
          <th>Stock</th>
          <th class="num">Value</th>
          <th class="num">ROI</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>
        {tech_action_rows()}
      </tbody>
    </table>
  </div>

  <!-- 52-Week Alerts -->
  <div class="card">
    <h2>52-Week Range Alerts</h2>
    <table>
      <thead>
        <tr>
          <th>Stock</th>
          <th class="num">Position in Range</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        {range_alert_rows()}
      </tbody>
    </table>
  </div>

</div>

<div class="card">
  <h2>Performance by Asset Type</h2>
  <small style="color:var(--muted)">Open positions only. Capital = price P&amp;L. Total Return includes dividends.</small>
  <table>
    <thead>
      <tr>
        <th>Asset Type</th>
        <th class="num">#</th>
        <th class="num">Invested</th>
        <th class="num">Value</th>
        <th class="num">Capital P&amp;L</th>
        <th class="num">Cap %</th>
        <th class="num">Divs</th>
        <th class="num">Total Return</th>
        <th class="num">ROI %</th>
      </tr>
    </thead>
    <tbody>{asset_bucket_pnl_rows()}</tbody>
  </table>
</div>

<!-- Analytics: Top Movers, Sector P&L, Concentration, Rebalancing -->
<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;" id="analytics">

<div class="card">
  <h2>Top Movers Today</h2>
  <table>
    <thead><tr><th>Stock</th><th class="num">Change</th><th class="num">Impact</th></tr></thead>
    <tbody>{top_movers_rows()}</tbody>
  </table>
</div>

<div class="card">
  <h2>Sector P&amp;L</h2>
  <table>
    <thead><tr><th>Sector</th><th class="num">Value</th><th class="num">P&amp;L</th><th class="num">ROI</th></tr></thead>
    <tbody>{sector_pnl_rows()}</tbody>
  </table>
</div>

<div class="card">
  <h2>Concentration Risk</h2>
  <small style="color:var(--muted)">UK stocks &gt;5% | Others &gt;10%</small>
  <table>
    <thead><tr><th>Stock</th><th class="num">Weight</th><th class="num">Value</th><th>Flag</th></tr></thead>
    <tbody>{concentration_rows()}</tbody>
  </table>
</div>

<div class="card">
  <h2>Rebalancing Actions</h2>
  <small style="color:var(--muted)">Triggered when &gt;2% off target</small>
  <table>
    <thead><tr><th>Asset Type</th><th class="num">Amount</th><th>Action</th></tr></thead>
    <tbody>{rebalance_rows()}</tbody>
  </table>
</div>

</div>

<!-- Asset Allocation -->
<div class="card">
  <h2>Asset Allocation vs Target</h2>
  <table>
    <thead>
      <tr>
        <th>Asset Type</th>
        <th class="num">Value</th>
        <th class="num">Actual %</th>
        <th class="num">Target %</th>
        <th class="num">Diff</th>
        <th class="num">Status</th>
      </tr>
    </thead>
    <tbody>
      {alloc_rows()}
    </tbody>
  </table>
</div>

<!-- News: Dividends & Earnings -->
<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;" id="news">
<div class="card">
  <h2>Upcoming Dividends (14 days)</h2>
  <table>
    <thead>
      <tr>
        <th>Stock</th>
        <th class="num">Ex-Date</th>
        <th class="num">Per Share</th>
        <th class="num">Yield</th>
        <th class="num">Est. Payment</th>
      </tr>
    </thead>
    <tbody>
      {dividend_rows()}
    </tbody>
  </table>
</div>
<div class="card">
  <h2>Upcoming Earnings (14 days)</h2>
  <table>
    <thead>
      <tr>
        <th>Stock</th>
        <th class="num">Date</th>
        <th class="num">Holding Value</th>
      </tr>
    </thead>
    <tbody>
      {earnings_rows()}
    </tbody>
  </table>
</div>
</div>

<!-- Portfolio Holdings -->
<div class="card" style="overflow-x:auto;" id="holdings">
  <h2>Portfolio Holdings ({len(holdings)} positions)</h2>
  <table>
    <thead>
      <tr>
        <th>Stock</th>
        <th>Sector</th>
        <th class="num">Invested</th>
        <th class="num">Value</th>
        <th class="num">P&amp;L</th>
        <th class="num">ROI</th>
        <th class="num">Divs</th>
        <th class="num">Yield</th>
        <th class="num">Price</th>
        <th class="num">VR</th>
        <th>Signal</th>
      </tr>
    </thead>
    <tbody>
      {holding_rows()}
    </tbody>
  </table>
</div>

<!-- Watchlist -->
<div class="card" id="watchlist">
  <h2>Watchlist ({len(watchlist)} stocks)</h2>
  <table>
    <thead>
      <tr>
        <th>Stock</th>
        <th>Sector</th>
        <th class="num">Live Price (GBP/USD)</th>
        <th class="num">Target (GBP/USD)</th>
        <th class="num">Yield</th>
        <th class="num">Value Ratio</th>
        <th>Signal</th>
      </tr>
    </thead>
    <tbody>
      {watchlist_rows()}
    </tbody>
  </table>
</div>

<footer>
  eToro Portfolio Dashboard &nbsp;|&nbsp; @Dalkent13 &nbsp;|&nbsp;
  Targets = avg(DCF, DDM) from Assumptions sheet &nbsp;|&nbsp;
  Signal: VR≥1.25 Strong Buy · ≥1.10 Buy · ≥0.90 Fair Value · ≥0.75 Sell · &lt;0.75 Strong Sell
</footer>

<script>
(function() {{
  var ts = parseInt(document.getElementById('gen-ts').getAttribute('data-ts'), 10) * 1000;
  var ageMin = (Date.now() - ts) / 60000;
  if (ageMin > 60) {{
    var hrs = (ageMin / 60).toFixed(1);
    document.getElementById('stale-banner').innerHTML =
      '<div class="stale-warning">⚠️ Prices are ' + hrs + ' hours old. Re-run generate_dashboard.py for fresh data.</div>';
  }}
}})();
</script>
{_THEME_JS}

</body>
</html>"""
    return html


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 56)
    print("  generate_dashboard.py")
    print("=" * 56)

    holdings, watchlist, assumptions, cash, gbpusd, closed = load_excel()
    realized_pnl = sum(c["pnl"] for c in closed)
    print(f"  Portfolio: {len(holdings)} holdings | Watchlist: {len(watchlist)} | Closed: {len(closed)} (realized ${realized_pnl:,.2f}) | Cash: ${cash:,.2f} | GBP/USD: {gbpusd}")

    all_items = holdings + watchlist
    prices, upcoming_divs, upcoming_earnings, range_data, div_income, daily_changes, benchmarks, perf_data = fetch_market_data(all_items, holdings)

    summary = enrich(holdings, watchlist, assumptions, cash, gbpusd, prices, closed)

    html = build_html(holdings, watchlist, summary, upcoming_divs, upcoming_earnings, range_data, div_income, daily_changes, benchmarks, perf_data, closed)
    OUTPUT.write_text(html, encoding="utf-8")

    print(f"\n  Dashboard saved -> {OUTPUT}")
    print(f"  Total value: {fmt_usd(summary['total_value'])}  |  Return: {fmt_pct(summary['total_roi'])}")
    print("  Open eToro_dashboard.html in your browser.\n")


if __name__ == "__main__":
    main()
