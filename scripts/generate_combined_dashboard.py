#!/usr/bin/env python3
"""
generate_combined_dashboard.py
──────────────────────────────
Builds combined_dashboard.html showing eToro + Trading 212 ISA side-by-side
with per-broker totals and a grand total.

Inputs:
    data/eToro_Master.xlsx         — Portfolio sheet + Assumptions (GBP/USD)
    data/t212_portfolio.json       — written by sync_t212.py

Output:
    combined_dashboard.html        — open in any browser
"""

import json
import sys
from datetime import datetime
from html import escape
from pathlib import Path

# Load etoro.env from project root before reading os.getenv anywhere below.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import _envloader  # noqa: F401  (side-effect import: populates os.environ)

import openpyxl

BASE_DIR  = Path(__file__).parent.parent
DATA_DIR  = BASE_DIR / "data"
MATCH_CSV = DATA_DIR / "etoro_portfolio_tickermatch.csv"
MASTER    = DATA_DIR / "eToro_Master.xlsx"
T212_JSON = DATA_DIR / "t212_portfolio.json"
T212_CSV  = DATA_DIR / "t212_portfolio_manual.csv"
T212_ACC  = DATA_DIR / "t212_account_manual.txt"
OUTPUT   = BASE_DIR / "dashboards" / "t212_dashboard.html"


# ── Formatting helpers ───────────────────────────────────────────────────────

def fmt_money(v, symbol="£", dp=2):
    if v is None:
        return "—"
    sign = "-" if v < 0 else ""
    return f"{sign}{symbol}{abs(v):,.{dp}f}"

def fmt_pct(v):
    if v is None:
        return "—"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.1f}%"

def cls_for(v):
    if v is None:
        return ""
    return "pos" if v >= 0 else "neg"


# ── eToro loader ─────────────────────────────────────────────────────────────

MASTER_JSON = DATA_DIR / "etoro_master.json"


def _to_float(v, default=0.0):
    if v in (None, "", "—"):
        return default
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return default


def load_etoro():
    """Load eToro state from data/etoro_master.json (produced by sync_xlsx_to_vault).

    Falls back to the xlsx directly only if the JSON cache is missing.
    """
    if not MASTER_JSON.exists():
        # Bootstrap fallback — run the sync once, then reload.
        if MASTER.exists():
            print(f"  WARNING: {MASTER_JSON.name} missing — run sync_xlsx_to_vault.py to build the cache")
        return {"holdings": [], "cash_usd": 0.0, "gbpusd": 1.27}

    data = json.loads(MASTER_JSON.read_text(encoding="utf-8"))
    sheets = data.get("sheets") or {}
    gbpusd = _to_float(data.get("assumptions", {}).get("GBP/USD"), default=1.27)

    holdings, cash_usd, open_divs_usd = [], 0.0, 0.0
    for obj in (sheets.get("portfolio") or {}).get("objects") or []:
        company = str(obj.get("Company Name") or "").strip()
        if not company:
            continue
        if company.upper() == "CASH":
            cash_usd = _to_float(obj.get("Invested (USD)"))
            continue
        if "GRAND TOTAL" in company.upper():
            continue
        position_divs = sum(
            _to_float(obj.get(col))
            for col in ("Div 2023 (USD)", "Div 2024 (USD)", "Div 2025 (USD)")
        )
        open_divs_usd += position_divs
        cv = _to_float(obj.get("Current Value (USD)"), default=0)
        holdings.append({
            "company":  company,
            "ticker":   str(obj.get("eToro Ticker") or "").strip(),
            "yahoo":    str(obj.get("Yahoo Ticker") or obj.get("eToro Ticker") or "").strip(),
            "currency": str(obj.get("Currency") or "USD").strip(),
            "units":    _to_float(obj.get("Units Held")),
            "invested_usd":      _to_float(obj.get("Invested (USD)")),
            "current_value_usd": cv if cv > 0 else None,
            "divs_usd":          position_divs,
        })

    closed_realised_usd, closed_divs_usd, closed_count = 0.0, 0.0, 0
    for obj in (sheets.get("closed_positions") or {}).get("objects") or []:
        invested_c = _to_float(obj.get("Invested (USD)"))
        sale_c     = _to_float(obj.get("Sale Value (USD)"))
        if invested_c == 0 and sale_c == 0:
            continue
        closed_realised_usd += sale_c - invested_c
        for col in ("Div 2023 (USD)", "Div 2024 (USD)", "Div 2025 (USD)"):
            closed_divs_usd += _to_float(obj.get(col))
        closed_count += 1

    return {
        "holdings":      holdings,
        "cash_usd":      cash_usd,
        "gbpusd":        gbpusd,
        "closed": {
            "realised_usd": closed_realised_usd,
            "divs_usd":     closed_divs_usd,
            "count":        closed_count,
        },
        "open_divs_usd": open_divs_usd,
    }


# ── eToro live prices (optional, via yfinance) ───────────────────────────────

def _load_etoro_assetid_to_ticker() -> dict:
    """Read etoro_portfolio_tickermatch.csv → {asset_id (int): etoro_ticker}."""
    import csv
    m = {}
    if not MATCH_CSV.exists():
        return m
    with open(MATCH_CSV, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            try:
                m[int(row["Asset_ID"])] = row["Ticker"].strip()
            except (KeyError, ValueError):
                continue
    return m


def fetch_etoro_api_prices() -> dict:
    """
    Call eToro's /trading/info/real/pnl and return {etoro_ticker: closeRate}.
    Returns {} if creds missing or call fails — caller can fall back to yfinance.
    """
    import os, uuid
    api_key  = os.getenv("ETORO_PUBLIC_API_KEY")
    user_key = os.getenv("ETORO_USER_KEY")
    if not api_key or not user_key:
        return {}

    try:
        import requests
        url = "https://public-api.etoro.com/api/v1/trading/info/real/pnl"
        headers = {"x-api-key": api_key, "x-user-key": user_key,
                   "x-request-id": str(uuid.uuid4()), "Accept": "application/json"}
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        print(f"  eToro API price fetch failed ({e}) — will fall back to yfinance")
        return {}

    id_to_ticker = _load_etoro_assetid_to_ticker()
    positions = data.get("clientPortfolio", {}).get("positions", [])
    prices = {}
    for pos in positions:
        iid = pos.get("instrumentID")
        if not isinstance(iid, int):
            continue
        ticker = id_to_ticker.get(iid)
        if not ticker:
            continue
        rate = (pos.get("unrealizedPnL") or {}).get("closeRate")
        if rate is not None:
            # Positions are fractional buys; later same-ticker entries would overwrite
            # with the same live price, which is fine.
            prices[ticker] = float(rate)
    print(f"  eToro API returned live prices for {len(prices)} tickers")
    return prices


def fetch_etoro_prices(holdings):
    """
    Live pricing. Primary source = eToro API (authoritative, matches app).
    Fallback = yfinance for any ticker eToro didn't return.
    Returns {yahoo_key_used_by_holdings: raw_price_local_ccy}.
    """
    # eToro prices are keyed by eToro ticker, holdings are keyed by "yahoo" (yahoo
    # ticker). The eToro ticker matches "yahoo" exactly for UK stocks (.L suffix)
    # and for most US tickers, but not always. Build both lookups.
    api_by_etoro = fetch_etoro_api_prices()
    prices = {}
    missing = []
    for h in holdings:
        et_ticker = h["ticker"]
        y_key     = h["yahoo"] or et_ticker
        rate = api_by_etoro.get(et_ticker) or api_by_etoro.get(y_key)
        if rate is not None:
            prices[y_key] = rate
        else:
            missing.append(h)

    if not missing:
        return prices

    # Fallback: yfinance for anything eToro didn't give us
    try:
        import yfinance as yf
    except ImportError:
        print(f"  yfinance not installed — {len(missing)} tickers will fall back to invested")
        return prices

    YF_OVERRIDES = {"BTC": "BTC-USD", "Roku": "ROKU"}
    targets = {h["yahoo"]: YF_OVERRIDES.get(h["yahoo"], h["yahoo"])
               for h in missing if h["yahoo"]}
    print(f"  Fetching {len(targets)} missing prices from yfinance ...")
    for orig, yf_t in targets.items():
        try:
            hist = yf.Ticker(yf_t).history(period="2d")
            if not hist.empty:
                prices[orig] = float(hist["Close"].iloc[-1])
        except Exception:
            pass
    return prices


def enrich_etoro(etoro, prices):
    """Attach current_value_usd to holdings and compute totals in GBP."""
    gbpusd = etoro["gbpusd"]
    holdings = etoro["holdings"]
    for h in holdings:
        raw = prices.get(h["yahoo"])
        if raw is not None:
            # convert to USD for consistency
            if h["currency"] == "GBp":
                price_usd = (raw / 100) * gbpusd
            elif h["currency"] == "GBP":
                price_usd = raw * gbpusd
            else:
                price_usd = raw
            h["current_value_usd"] = h["units"] * price_usd
        elif h["current_value_usd"] is None:
            h["current_value_usd"] = h["invested_usd"]

        h["pnl_usd"] = h["current_value_usd"] - h["invested_usd"]
        h["roi"]     = (h["pnl_usd"] / h["invested_usd"] * 100) if h["invested_usd"] else 0

        # GBP equivalents for combined totals
        h["invested_gbp"]      = h["invested_usd"] / gbpusd
        h["current_value_gbp"] = h["current_value_usd"] / gbpusd
        h["pnl_gbp"]           = h["pnl_usd"] / gbpusd

    holdings.sort(key=lambda x: x["current_value_usd"], reverse=True)

    invested_usd = sum(h["invested_usd"] for h in holdings) + etoro["cash_usd"]
    value_usd    = sum(h["current_value_usd"] for h in holdings) + etoro["cash_usd"]
    pnl_usd      = value_usd - invested_usd

    # Lifetime return: open capital P&L + realised P&L on closed trades + all dividends
    closed        = etoro.get("closed", {}) or {}
    realised_usd  = closed.get("realised_usd", 0) or 0
    closed_divs   = closed.get("divs_usd", 0) or 0
    open_divs     = etoro.get("open_divs_usd", 0) or 0
    total_divs    = open_divs + closed_divs
    lifetime_usd  = pnl_usd + realised_usd + total_divs

    etoro["totals"] = {
        "invested_usd":  invested_usd,
        "value_usd":     value_usd,
        "pnl_usd":       pnl_usd,
        "invested_gbp":  invested_usd / gbpusd,
        "value_gbp":     value_usd / gbpusd,
        "pnl_gbp":       pnl_usd / gbpusd,
        "cash_gbp":      etoro["cash_usd"] / gbpusd,
        "roi":           (pnl_usd / invested_usd * 100) if invested_usd else 0,
        "realised_gbp":  realised_usd / gbpusd,
        "divs_gbp":      total_divs  / gbpusd,
        "open_divs_gbp":   open_divs   / gbpusd,
        "closed_divs_gbp": closed_divs / gbpusd,
        "lifetime_gbp":  lifetime_usd / gbpusd,
        "closed_count":  closed.get("count", 0),
    }
    return etoro


# ── T212 loader ──────────────────────────────────────────────────────────────

def _load_t212_from_csv():
    """Fallback when the API is unavailable — read manual CSV + account file."""
    import csv
    if not T212_CSV.exists():
        return None

    positions = []
    with open(T212_CSV, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or line.lower().startswith("ticker,"):
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 5:
                continue
            try:
                positions.append({
                    "ticker":        parts[0],
                    "raw_ticker":    parts[0],
                    "quantity":      float(parts[1]),
                    "average_price": float(parts[2]),
                    "current_price": float(parts[3]),
                    "ppl":           float(parts[4]),
                    "fx_ppl":        None,
                    "initial_fill":  None,
                    "pie_quantity":  0,
                })
            except ValueError:
                continue

    cash = {"free": 0, "invested": 0, "total": 0, "ppl": 0, "result": 0, "pie_cash": 0}
    if T212_ACC.exists():
        for raw in T212_ACC.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            k = k.strip().lower()
            try:
                v = float(v.strip())
            except ValueError:
                continue
            if   k == "free_cash": cash["free"] = v
            elif k == "invested":  cash["invested"] = v
            elif k == "total":     cash["total"] = v
            elif k == "ppl":       cash["ppl"] = v

    return {
        "generated_at":     datetime.now().isoformat(timespec="seconds"),
        "account_currency": "GBP",
        "account_id":       None,
        "cash":             cash,
        "positions":        positions,
        "_source":          "manual CSV",
    }


def load_t212():
    if T212_JSON.exists():
        data = json.loads(T212_JSON.read_text(encoding="utf-8"))
        data.setdefault("_source", "API")
    else:
        data = _load_t212_from_csv()
        if data is None:
            print(f"  WARNING: no T212 data — run sync_t212.py or fill in {T212_CSV.name}")
            return None
        print(f"  T212 API data missing — using manual CSV fallback "
              f"({len(data['positions'])} positions)")

    # Positions: value in instrument ccy (approx). P&L is already account ccy (GBP).
    for p in data["positions"]:
        p["value_instr"] = p["quantity"] * p["current_price"]
        p["cost_instr"]  = p["quantity"] * p["average_price"]
        p["roi"] = (p["ppl"] / p["cost_instr"] * 100) if p["cost_instr"] else 0

    data["positions"].sort(key=lambda x: x["value_instr"], reverse=True)

    cash = data["cash"]
    # Account-level totals come straight from /equity/account/cash
    invested_gbp = cash.get("invested", 0) or 0
    positions_gbp = (cash.get("total", 0) or 0) - (cash.get("free", 0) or 0)
    ppl_gbp       = cash.get("ppl", 0) or 0
    total_gbp     = cash.get("total", 0) or 0
    free_gbp      = cash.get("free", 0) or 0

    # Lifetime return — money-weighted: total value + withdrawals − deposits, over deposits
    life = data.get("lifetime") or {}
    deposits    = life.get("deposits", 0) or 0
    withdrawals = life.get("withdrawals", 0) or 0
    dividends   = life.get("dividends", 0) or 0
    net_deposited = max(deposits - withdrawals, 0.0001)  # guard div-zero
    lifetime_return = total_gbp + withdrawals - deposits  # includes realised + unrealised + dividends
    lifetime_roi    = (lifetime_return / net_deposited * 100) if deposits else None

    # "App-style" ROI: mirrors the percentage T212 shows in its UI, using their
    # own `result` field over their `invested` cost basis. Our £ figure already
    # matches the app within rounding; this just gives a % that matches too.
    app_style_return = cash.get("result")
    app_style_roi = None
    if app_style_return and invested_gbp:
        app_style_roi = app_style_return / invested_gbp * 100

    data["totals"] = {
        "invested_gbp":      invested_gbp,
        "positions_gbp":     positions_gbp,
        "free_gbp":          free_gbp,
        "total_gbp":         total_gbp,
        "ppl_gbp":           ppl_gbp,
        "roi":               (ppl_gbp / invested_gbp * 100) if invested_gbp else 0,
        "deposits":          deposits,
        "withdrawals":       withdrawals,
        "dividends":         dividends,
        "lifetime_return":   lifetime_return,
        "lifetime_roi":      lifetime_roi,
        "app_style_return":  app_style_return,
        "app_style_roi":     app_style_roi,
    }
    return data


# ── HTML ─────────────────────────────────────────────────────────────────────

try:
    from editorial_theme import CSS as _EDITORIAL_CSS, FONTS_LINK as _FONTS_LINK, nav_html as _nav_html, THEME_JS as _THEME_JS
except ImportError:
    from scripts.editorial_theme import CSS as _EDITORIAL_CSS, FONTS_LINK as _FONTS_LINK, nav_html as _nav_html, THEME_JS as _THEME_JS

_LEGACY_CSS_UNUSED = """
* { box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    margin: 0; padding: 24px;
    background: #0f172a; color: #e2e8f0;
}
h1 { margin: 0 0 4px 0; font-size: 22px; }
.sub { color: #94a3b8; font-size: 13px; margin-bottom: 24px; }
.grid {
    display: grid; gap: 16px;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    margin-bottom: 28px;
}
.card {
    background: #1e293b; border: 1px solid #334155;
    border-radius: 10px; padding: 16px;
}
.card .label { color: #94a3b8; font-size: 11px; text-transform: uppercase; letter-spacing: .06em; }
.card .value { font-size: 22px; font-weight: 600; margin-top: 4px; }
.card .sub2  { font-size: 12px; color: #94a3b8; margin-top: 4px; }
.card.grand  { background: linear-gradient(135deg, #1e3a8a, #0f766e); border-color: #0ea5e9; }
.section {
    background: #1e293b; border: 1px solid #334155;
    border-radius: 10px; padding: 18px; margin-bottom: 24px;
}
.section h2 { margin: 0 0 4px 0; font-size: 17px; }
.section .meta { color: #94a3b8; font-size: 12px; margin-bottom: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { padding: 8px 10px; text-align: left; border-bottom: 1px solid #334155; }
th { color: #94a3b8; font-weight: 500; font-size: 11px; text-transform: uppercase; letter-spacing: .05em; }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
tbody tr:hover { background: #283548; }
.pos { color: #10b981; }
.neg { color: #ef4444; }
.broker-bar { display: flex; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; }
.chip { background: #0f172a; border: 1px solid #334155; border-radius: 999px;
        padding: 4px 12px; font-size: 12px; }
.footer { color: #64748b; font-size: 11px; margin-top: 32px; text-align: center; }
"""


def build_html(etoro, t212):
    g = etoro["gbpusd"]
    et_totals = etoro["totals"]

    # Grand totals in GBP
    grand_value    = et_totals["value_gbp"]    + (t212["totals"]["total_gbp"] if t212 else 0)
    grand_invested = et_totals["invested_gbp"] + (t212["totals"]["invested_gbp"] if t212 else 0)
    grand_pnl      = grand_value - grand_invested
    grand_roi      = (grand_pnl / grand_invested * 100) if grand_invested else 0

    # Grand LIFETIME return — full picture across both brokers:
    #   eToro: open P&L + closed realised P&L + all dividends ever received
    #   T212:  money-weighted lifetime return (total + withdrawals − deposits)
    t212_lt_return = (t212["totals"].get("lifetime_return") if t212 else None) or 0
    grand_lifetime_return = et_totals["lifetime_gbp"] + t212_lt_return

    # ── KPI hero cards ──
    cards = []
    if t212:
        tt = t212["totals"]
        lifetime_return = tt.get("lifetime_return") or 0
        lifetime_roi    = tt.get("lifetime_roi")
        app_roi         = tt.get("app_style_roi")
        dividends       = tt.get("dividends") or 0
        deposits        = tt.get("deposits") or 0
        withdrawals     = tt.get("withdrawals") or 0
        net_deposited   = deposits - withdrawals
        raw = t212 or {}
        lt = raw.get("lifetime") or {}
        dividend_count  = lt.get("dividend_count") or 0
        txn_count       = lt.get("transaction_count") or 0

        cards.append(f"""
      <div class="card grand">
        <div class="label">Total Value</div>
        <div class="value">{fmt_money(tt["total_gbp"])}</div>
        <div class="sub2">{len(t212["positions"])} positions · free £{tt["free_gbp"]:,.2f}</div>
      </div>""")

        lt_detail = (f'{fmt_pct(lifetime_roi)}' + (f" · app {fmt_pct(app_roi)}" if app_roi is not None else "")) if lifetime_roi is not None else ""
        cards.append(f"""
      <div class="card">
        <div class="label">Lifetime Return</div>
        <div class="value {cls_for(lifetime_return)}">{fmt_money(lifetime_return)}</div>
        <div class="sub2">{lt_detail}</div>
      </div>""")

        cards.append(f"""
      <div class="card">
        <div class="label">Open P&amp;L</div>
        <div class="value {cls_for(tt["ppl_gbp"])}">{fmt_money(tt["ppl_gbp"])}</div>
        <div class="sub2">{fmt_pct(tt["roi"])} · vs £{tt["invested_gbp"]:,.0f} invested</div>
      </div>""")

        cards.append(f"""
      <div class="card">
        <div class="label">Dividends</div>
        <div class="value pos">{fmt_money(dividends)}</div>
        <div class="sub2">{dividend_count} payments</div>
      </div>""")

        cards.append(f"""
      <div class="card">
        <div class="label">Net Deposited</div>
        <div class="value">{fmt_money(net_deposited)}</div>
        <div class="sub2">in £{deposits:,.0f} · out £{withdrawals:,.0f}</div>
      </div>""")

        cards.append(f"""
      <div class="card">
        <div class="label">Activity</div>
        <div class="value">{txn_count}</div>
        <div class="sub2">total transactions</div>
      </div>""")
    else:
        cards.append("""
      <div class="card">
        <div class="label">Trading 212 ISA</div>
        <div class="value">—</div>
        <div class="sub2">Run sync_t212.py</div>
      </div>""")

    # ── eToro table ──
    et_rows = []
    for h in etoro["holdings"]:
        et_rows.append(f"""
          <tr>
            <td>{escape(h["company"])}</td>
            <td>{escape(h["ticker"])}</td>
            <td class="num units-col">{h["units"]:,.4f}</td>
            <td class="num">{fmt_money(h["invested_gbp"])}</td>
            <td class="num">{fmt_money(h["current_value_gbp"])}</td>
            <td class="num {cls_for(h["pnl_gbp"])}">{fmt_money(h["pnl_gbp"])}</td>
            <td class="num {cls_for(h["roi"])}">{fmt_pct(h["roi"])}</td>
          </tr>""")

    etoro_section = f"""
    <div class="section">
      <h2>eToro</h2>
      <div class="meta">
        {len(etoro["holdings"])} open · {et_totals["closed_count"]} closed
        (realised {fmt_money(et_totals["realised_gbp"])}) ·
        dividends {fmt_money(et_totals["divs_gbp"])} ·
        cash {fmt_money(et_totals["cash_gbp"])} ·
        invested {fmt_money(et_totals["invested_gbp"])} ·
        GBP/USD {g:.4f}
      </div>
      <table>
        <thead>
          <tr>
            <th>Company</th><th>Ticker</th>
            <th class="num units-col">Units</th>
            <th class="num">Invested (GBP)</th>
            <th class="num">Value (GBP)</th>
            <th class="num">P&amp;L (GBP)</th>
            <th class="num">ROI</th>
          </tr>
        </thead>
        <tbody>{''.join(et_rows) or '<tr><td colspan="7">No positions.</td></tr>'}</tbody>
      </table>
    </div>"""

    # ── T212 table ──
    if t212:
        # Compute GBP value per position (best-effort based on instrument currency).
        from datetime import datetime as _dt, timezone as _tz
        now_dt = _dt.now(_tz.utc)
        gbpusd = etoro.get("gbpusd") or 1.0

        def _value_gbp(p: dict) -> float | None:
            qty   = p.get("quantity") or 0
            price = p.get("current_price") or 0
            ccy   = (p.get("instrument_ccy") or "").upper()
            if ccy == "GBX":  return qty * price / 100
            if ccy == "GBP":  return qty * price
            if ccy == "USD":  return qty * price / gbpusd
            return None  # unknown ccy — skip GBP estimate

        def _hold_days(p: dict) -> int | None:
            iso = p.get("initial_fill")
            if not iso:
                return None
            try:
                dt = _dt.fromisoformat(iso.replace("Z", "+00:00"))
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=_tz.utc)
                return max(0, (now_dt - dt).days)
            except (ValueError, AttributeError):
                return None

        # Precompute values and total for weight %
        enriched = [(p, _value_gbp(p), _hold_days(p)) for p in t212["positions"]]
        total_val = sum(v for _, v, _ in enriched if v) or 1
        enriched.sort(key=lambda x: x[1] or 0, reverse=True)

        t_rows = []
        for p, val_gbp, hold_days in enriched:
            name = p.get("name") or p["ticker"]
            ccy  = p.get("instrument_ccy", "")
            ccy_badge = f' <span style="color:var(--ink-3);font-size:10px;font-family:var(--f-mono);text-transform:uppercase;letter-spacing:.05em;">{escape(ccy)}</span>' if ccy else ""
            val_str = fmt_money(val_gbp) if val_gbp else "—"
            weight_str = f'{val_gbp / total_val * 100:.1f}%' if val_gbp else "—"
            hold_str = ""
            if hold_days is not None:
                if hold_days < 60:
                    hold_str = f'{hold_days}d'
                elif hold_days < 365:
                    hold_str = f'{hold_days // 30}mo'
                else:
                    hold_str = f'{hold_days // 365}y {(hold_days % 365) // 30}mo'
            t_rows.append(f"""
              <tr>
                <td>{escape(name)}{ccy_badge}</td>
                <td>{escape(p["ticker"])}</td>
                <td class="num units-col">{p["quantity"]:,.4f}</td>
                <td class="num">{val_str}</td>
                <td class="num">{weight_str}</td>
                <td class="num {cls_for(p["ppl"])}">{fmt_money(p["ppl"])}</td>
                <td class="num {cls_for(p["roi"])}">{fmt_pct(p["roi"])}</td>
                <td class="num" style="color:var(--ink-3);font-size:11px">{hold_str}</td>
              </tr>""")

        # Activity breakdown panel
        lt = t212.get("lifetime") or {}
        type_breakdown = lt.get("type_breakdown") or {}
        activity_rows = []
        for kind in ("DEPOSIT", "TRANSFER", "WITHDRAW"):
            if kind not in type_breakdown:
                continue
            row = type_breakdown[kind]
            activity_rows.append(
                f'<tr><td>{kind.title()}</td>'
                f'<td class="num">{row.get("count", 0)}</td>'
                f'<td class="num">{fmt_money(row.get("sum", 0))}</td></tr>'
            )
        if lt.get("dividend_count"):
            activity_rows.append(
                f'<tr><td>Dividends</td>'
                f'<td class="num">{lt["dividend_count"]}</td>'
                f'<td class="num pos">{fmt_money(lt.get("dividends", 0))}</td></tr>'
            )
        activity_table = (
            '<div class="section">'
            '<h2>Activity</h2>'
            '<table><thead><tr><th>Type</th><th class="num">Count</th><th class="num">Total (GBP)</th></tr></thead>'
            f'<tbody>{"".join(activity_rows) or "<tr><td colspan=3>No activity.</td></tr>"}</tbody></table>'
            '</div>'
        ) if activity_rows else ''

        tt = t212["totals"]
        holdings_section = f"""
        <div class="section">
          <h2>Holdings</h2>
          <div class="meta">
            {len(t212["positions"])} positions · free £{tt["free_gbp"]:,.2f} ·
            invested £{tt["invested_gbp"]:,.2f} · ccy {escape(t212["account_currency"])}
          </div>
          <table>
            <thead>
              <tr>
                <th>Company</th>
                <th>Ticker</th>
                <th class="num units-col">Units</th>
                <th class="num">Value (GBP)</th>
                <th class="num">Weight</th>
                <th class="num">P&amp;L (GBP)</th>
                <th class="num">ROI</th>
                <th class="num">Held</th>
              </tr>
            </thead>
            <tbody>{''.join(t_rows) or '<tr><td colspan="8">No positions.</td></tr>'}</tbody>
          </table>
        </div>"""
        t212_section = holdings_section + activity_table
    else:
        t212_section = """
        <div class="section">
          <h2>Trading 212 ISA</h2>
          <div class="meta">No data yet. Configure t212.env and run sync_t212.py.</div>
        </div>"""

    now = datetime.now().strftime("%d %b %Y %H:%M")
    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Trading 212 ISA</title>
{_FONTS_LINK}
<style>{_EDITORIAL_CSS}</style>
</head><body>
{_nav_html(active="t212", privacy=True)}
<h1>Trading 212 ISA</h1>
<div class="sub">Generated {now}</div>
<div class="grid-cards">{''.join(cards)}</div>
{t212_section}
<div class="footer">
  T212 instrument prices shown in their native currency (GBX = pence for UK stocks).
</div>
{_THEME_JS}
</body></html>"""


# ── Main ─────────────────────────────────────────────────────────────────────

def write_combined_json(etoro, t212, gbpusd):
    """Emit data/combined_portfolio.json so other dashboards can consume
    the combined view without re-loading xlsx/JSON sources."""
    holdings = []
    for h in etoro["holdings"]:
        holdings.append({
            "broker": "eToro",
            "ticker": h.get("ticker") or "",
            "yahoo":  h.get("yahoo") or h.get("ticker") or "",
            "name":   h.get("company") or "",
            "currency": h.get("currency") or "USD",
            "units":  h.get("units") or 0,
            "value_gbp": h.get("current_value_gbp") or 0,
            "pnl_gbp":   h.get("pnl_gbp") or 0,
            "roi":       h.get("roi") or 0,
        })
    if t212:
        for p in t212.get("positions", []):
            # Value/cost in T212 are in instrument currency; P&L is GBP.
            # Use current_price × qty × fx (approx) for GBP value.
            ticker_raw = p.get("ticker") or p.get("raw_ticker") or ""
            # T212 ticker uses suffixes like 'l' (LSE), 'd' (Xetra), 'p' (Amsterdam).
            # Map LSE: OCDOl -> OCDO.L for Yahoo compatibility.
            yahoo = ticker_raw
            if ticker_raw.endswith("l") and ticker_raw[:-1].isupper():
                yahoo = ticker_raw[:-1] + ".L"
            value_instr = p.get("value_instr") or 0
            ccy = (p.get("instrument_ccy") or "GBP").upper()
            if ccy == "GBX" or ccy == "GBP":
                value_gbp = value_instr / 100 if ccy == "GBX" else value_instr
            elif ccy == "USD":
                value_gbp = value_instr / gbpusd
            else:
                value_gbp = value_instr
            holdings.append({
                "broker": "T212",
                "ticker": ticker_raw,
                "yahoo":  yahoo,
                "name":   p.get("name") or ticker_raw,
                "currency": ccy,
                "units":  p.get("quantity") or 0,
                "value_gbp": value_gbp,
                "pnl_gbp":   p.get("ppl") or 0,
                "roi":       p.get("roi") or 0,
            })
    total_value_gbp = sum(h["value_gbp"] for h in holdings)
    total_pnl_gbp   = sum(h["pnl_gbp"]   for h in holdings)
    for h in holdings:
        h["weight"] = (h["value_gbp"] / total_value_gbp * 100) if total_value_gbp else 0

    et_t = etoro.get("totals", {}) or {}
    t2_t = (t212 or {}).get("totals", {}) or {}
    data = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "gbpusd": gbpusd,
        "total_value_gbp": total_value_gbp,
        "total_pnl_gbp":   total_pnl_gbp,
        "positions":       len(holdings),
        "by_broker": {
            "etoro": {
                "value_gbp":     et_t.get("value_gbp", 0),
                "pnl_gbp":       et_t.get("pnl_gbp", 0),
                "cash_gbp":      et_t.get("cash_gbp", 0),
                "positions":     len(etoro.get("holdings") or []),
            },
            "t212": {
                "value_gbp":     t2_t.get("total_gbp", 0),
                "pnl_gbp":       t2_t.get("ppl_gbp", 0),
                "cash_gbp":      t2_t.get("free_gbp", 0),
                "positions":     len((t212 or {}).get("positions") or []),
            },
        },
        "holdings": holdings,
    }
    out = DATA_DIR / "combined_portfolio.json"
    out.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"  Wrote {out}")


def main():
    print("  Loading eToro Master ...")
    etoro = load_etoro()

    prices = fetch_etoro_prices(etoro["holdings"])
    enrich_etoro(etoro, prices)

    print("  Loading Trading 212 JSON ...")
    t212 = load_t212()

    write_combined_json(etoro, t212, etoro["gbpusd"])

    html = build_html(etoro, t212)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"  Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
