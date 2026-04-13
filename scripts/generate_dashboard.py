#!/usr/bin/env python3
"""
generate_dashboard.py
─────────────────────
Generates a self-contained HTML portfolio dashboard from eToro_Master.xlsx.
Fetches live prices via yfinance and computes P&L, ROI, and valuation signals.

Usage:
    python scripts/generate_dashboard.py

Output:
    dashboard.html — open directly in any browser. No server required.
"""

import sys
import json
from pathlib import Path
from datetime import datetime

import openpyxl

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
MASTER   = DATA_DIR / "eToro_Master.xlsx"
OUTPUT   = BASE_DIR / "dashboard.html"


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


# ── Load Excel ─────────────────────────────────────────────────────────────────

def load_excel():
    print(f"  Reading {MASTER} ...")
    wb = openpyxl.load_workbook(str(MASTER), data_only=True)

    # GBP/USD
    ws_a = wb["Assumptions"]
    gbpusd = 1.34
    for row in ws_a.iter_rows(min_row=3, max_row=6, values_only=True):
        if row[0] == "GBP/USD" and row[1]:
            try:
                gbpusd = float(row[1])
            except Exception:
                pass
            break

    # Assumptions: blended target per ticker (col L, index 11)
    assumptions = {}
    for row in ws_a.iter_rows(min_row=7, max_row=300, values_only=True):
        if not row[0]:
            continue
        ticker = str(row[0]).strip()
        blended = None
        try:
            if row[11] is not None:
                blended = float(row[11])
        except (TypeError, ValueError):
            pass
        if blended is not None:
            assumptions[ticker] = {"blended": blended}

    # Portfolio holdings
    ws_p = wb["Portfolio"]
    holdings = []
    cash_balance = 0.0
    for row in ws_p.iter_rows(min_row=3, max_row=300, values_only=True):
        company = str(row[1] or "").strip()
        if not company:
            continue
        if company == "CASH":
            try:
                cash_balance = float(row[10] or 0)
            except Exception:
                pass
            continue
        if "GRAND TOTAL" in company.upper():
            continue

        ticker   = str(row[3] or "").strip()
        yahoo    = str(row[4] or ticker).strip()
        sector   = str(row[2] or "Other").strip()
        currency = str(row[5] or "USD").strip()

        try:
            units    = float(row[8]  or 0)
            invested = float(row[10] or 0)
        except Exception:
            continue

        divs = sum(
            float(row[i] or 0)
            for i in (17, 18, 19, 20)
        )

        holdings.append({
            "company":   company,
            "ticker":    ticker,
            "yahoo":     yahoo,
            "sector":    sector,
            "currency":  currency,
            "units":     units,
            "invested":  invested,
            "total_divs": divs,
        })

    # Ticker → company name lookup from Tickers sheet
    ticker_names = {}
    if "Tickers" in wb.sheetnames:
        ws_t = wb["Tickers"]
        for row in ws_t.iter_rows(min_row=3, max_row=500, values_only=True):
            if row[1] and row[2]:
                ticker_names[str(row[2]).strip()] = str(row[1]).strip()  # FTSE ticker
            if row[1] and row[5]:
                ticker_names[str(row[5]).strip()] = str(row[1]).strip()  # Yahoo ticker

    # Watchlist
    ws_w = wb["Watchlist"]
    watchlist = []
    for row in ws_w.iter_rows(min_row=3, max_row=300, values_only=True):
        ticker   = str(row[3] or "").strip()
        if not ticker:
            continue
        yahoo    = str(row[4] or ticker).strip()
        company  = (str(row[1]).strip() if row[1]
                    else ticker_names.get(ticker)
                    or ticker_names.get(yahoo)
                    or ticker)
        sector   = str(row[2] or "").strip()
        currency = str(row[5] or "GBp").strip()
        if not ticker:
            continue
        watchlist.append({
            "company":  company,
            "ticker":   ticker,
            "yahoo":    yahoo,
            "sector":   sector,
            "currency": currency,
        })

    return holdings, watchlist, assumptions, cash_balance, gbpusd


# ── Fetch live prices ─────────────────────────────────────────────────────────

def fetch_prices(items):
    """Fetch current prices via yfinance. Returns {yahoo_ticker: price_in_local_currency}."""
    try:
        import yfinance as yf
    except ImportError:
        print("  Warning: yfinance not installed — run: pip install yfinance")
        return {}

    # Build ticker list with yfinance overrides
    YF_OVERRIDES = {"BTC": "BTC-USD", "Roku": "ROKU"}
    raw_to_yf = {}
    for i in items:
        y = i.get("yahoo", "")
        if y:
            raw_to_yf[y] = YF_OVERRIDES.get(y, y)

    yf_tickers = list(set(raw_to_yf.values()))
    if not yf_tickers:
        return {}

    prices = {}  # keyed by ORIGINAL yahoo ticker
    print(f"  Fetching {len(yf_tickers)} prices from Yahoo Finance ...")

    # Build a lookup of currency per original ticker
    currency_map = {i.get("yahoo", ""): i.get("currency", "") for i in items}

    # Fetch individually — more reliable than batch across UK + US mixed lists
    for orig, yf_t in raw_to_yf.items():
        try:
            t_obj = yf.Ticker(yf_t)
            hist = t_obj.history(period="2d")
            if not hist.empty:
                price = float(hist["Close"].iloc[-1])
                # Yahoo inconsistently returns some GBp stocks in GBP (100x too small).
                # Detect this by checking if Yahoo's reported currency is NOT GBp/GBX
                # for a stock we know should be in pence.
                if currency_map.get(orig) == "GBp":
                    yf_currency = (t_obj.info.get("currency") or "").upper()
                    # If Yahoo returns a non-GBP currency (USD, EUR etc.) for a stock
                    # we know is GBp, the price is in the wrong unit — multiply by 100
                    if yf_currency not in ("GBP", "GBX", "GBP", ""):
                        price = price * 100
                prices[orig] = price
        except Exception:
            pass

    return prices


# ── Compute derived metrics ───────────────────────────────────────────────────

def enrich(holdings, watchlist, assumptions, cash, gbpusd, prices):
    """Attach live_price, current_value, pnl, roi, target, vr, signal to each holding."""

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

    capital_pnl  = total_value - total_invested
    total_return = capital_pnl + total_divs
    total_roi    = (total_return / total_invested * 100) if total_invested else 0

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
        "total_invested":  total_invested,
        "total_value":     total_value,
        "cash":            cash,
        "capital_pnl":     capital_pnl,
        "total_divs":      total_divs,
        "total_return":    total_return,
        "total_roi":       total_roi,
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

def build_html(holdings, watchlist, summary):

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
              <td class="num">{price_str}</td>
              <td class="num">{vr_str}</td>
              <td>{signal_badge}</td>
            </tr>""")
        return "\n".join(rows)

    def watchlist_rows():
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
            rows.append(f"""
            <tr>
              <td><strong>{w["ticker"]}</strong><br><small>{w["company"][:30]}</small></td>
              <td><small>{w["sector"] or "—"}</small></td>
              <td class="num">{w.get("live_price_display","—")}</td>
              <td class="num">{target_str}</td>
              <td class="num">{fmt_vr(w["vr"])}</td>
              <td>{signal_badge}</td>
            </tr>""")
        return "\n".join(rows)

    pnl_colour  = "#10b981" if summary["capital_pnl"] >= 0 else "#ef4444"
    roi_colour  = "#10b981" if summary["total_roi"]   >= 0 else "#ef4444"
    live_note   = "Live prices via Yahoo Finance" if summary["prices_live"] else "⚠️ Prices unavailable — install yfinance"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>eToro Portfolio Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  :root {{
    --bg:      #0f172a;
    --card:    #1e293b;
    --border:  #334155;
    --text:    #e2e8f0;
    --muted:   #94a3b8;
    --accent:  #3b82f6;
    --pos:     #10b981;
    --neg:     #ef4444;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: var(--bg); color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    font-size: 13px; padding: 16px;
  }}
  h1 {{ font-size: 1.4rem; font-weight: 700; color: #fff; }}
  h2 {{ font-size: 1rem; font-weight: 600; color: var(--muted); text-transform: uppercase;
        letter-spacing: 0.05em; margin-bottom: 14px; }}
  .header {{
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 24px;
  }}
  .header-right {{ text-align: right; color: var(--muted); font-size: 0.8rem; }}
  .kpi-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 14px; margin-bottom: 28px;
  }}
  .kpi {{
    background: var(--card); border: 1px solid var(--border);
    border-radius: 10px; padding: 16px 20px;
  }}
  .kpi-label {{ font-size: 0.75rem; color: var(--muted); text-transform: uppercase;
                letter-spacing: 0.04em; margin-bottom: 6px; }}
  .kpi-value {{ font-size: 1.35rem; font-weight: 700; }}
  .kpi-sub   {{ font-size: 0.78rem; color: var(--muted); margin-top: 3px; }}
  .card {{
    background: var(--card); border: 1px solid var(--border);
    border-radius: 10px; padding: 14px 16px; margin-bottom: 16px; overflow: hidden;
  }}
  .signal-grid {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; }}
  table {{ width: 100%; border-collapse: collapse; }}
  th {{
    text-align: left; font-size: 0.68rem; text-transform: uppercase;
    letter-spacing: 0.04em; color: var(--muted); padding: 6px 8px;
    border-bottom: 1px solid var(--border); white-space: nowrap;
  }}
  td {{
    padding: 6px 8px; border-bottom: 1px solid rgba(51,65,85,0.5);
    vertical-align: middle; line-height: 1.3;
  }}
  td small {{ color: var(--muted); font-size: 0.72rem; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: rgba(255,255,255,0.03); }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .pos {{ color: var(--pos); }}
  .neg {{ color: var(--neg); }}
  .badge {{
    display: inline-block; padding: 3px 9px; border-radius: 20px;
    font-size: 0.72rem; font-weight: 600; color: #fff; white-space: nowrap;
  }}
  .stale {{ color: var(--muted); font-style: italic; }}
  .chart-wrap {{ display: flex; align-items: center; justify-content: center; height: 290px; }}
  footer {{
    margin-top: 30px; text-align: center;
    color: var(--muted); font-size: 0.75rem;
  }}
  @media (max-width: 900px) {{
    .signal-grid {{ grid-template-columns: 1fr; }}
  }}
</style>
</head>
<body>

<div class="header">
  <h1>eToro Portfolio Dashboard</h1>
  <div class="header-right">
    Generated: {summary["generated_at"]}<br>
    {live_note}
  </div>
</div>

<!-- KPI Summary -->
<div class="kpi-grid">
  <div class="kpi">
    <div class="kpi-label">Total Invested</div>
    <div class="kpi-value">{fmt_usd(summary["total_invested"])}</div>
    <div class="kpi-sub">incl. cash</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Portfolio Value</div>
    <div class="kpi-value">{fmt_usd(summary["total_value"])}</div>
    <div class="kpi-sub">Cash: {fmt_usd(summary["cash"])}</div>
  </div>
  <div class="kpi">
    <div class="kpi-label">Capital P&amp;L</div>
    <div class="kpi-value" style="color:{pnl_colour}">{fmt_usd(summary["capital_pnl"])}</div>
    <div class="kpi-sub">vs. invested</div>
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

<!-- Action Signals: 3 tables side by side -->
<div class="signal-grid">

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

<!-- Portfolio Holdings -->
<div class="card" style="overflow-x:auto;">
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
<div class="card">
  <h2>Watchlist ({len(watchlist)} stocks)</h2>
  <table>
    <thead>
      <tr>
        <th>Stock</th>
        <th>Sector</th>
        <th class="num">Live Price (GBP/USD)</th>
        <th class="num">Target (GBP/USD)</th>
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


</body>
</html>"""
    return html


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 56)
    print("  generate_dashboard.py")
    print("=" * 56)

    holdings, watchlist, assumptions, cash, gbpusd = load_excel()
    print(f"  Portfolio: {len(holdings)} holdings | Watchlist: {len(watchlist)} | Cash: ${cash:,.2f} | GBP/USD: {gbpusd}")

    all_items = holdings + watchlist
    prices = fetch_prices(all_items)
    print(f"  Prices fetched: {len(prices)} / {len(all_items)}")

    summary = enrich(holdings, watchlist, assumptions, cash, gbpusd, prices)

    html = build_html(holdings, watchlist, summary)
    OUTPUT.write_text(html, encoding="utf-8")

    print(f"\n  Dashboard saved → {OUTPUT}")
    print(f"  Total value: {fmt_usd(summary['total_value'])}  |  Return: {fmt_pct(summary['total_roi'])}")
    print("  Open dashboard.html in your browser.\n")


if __name__ == "__main__":
    main()
