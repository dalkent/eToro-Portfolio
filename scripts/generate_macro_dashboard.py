#!/usr/bin/env python3
"""
generate_macro_dashboard.py
───────────────────────────
Reads data/macro.json and writes macro_dashboard.html.

Sections:
  • Headline cards (UK 10Y, US 10Y, DXY, FTSE, S&P)
  • Rates & Yields
  • FX
  • Equity markets
  • Economic indicators (GDP / CPI / Unemployment / Central bank rates)
  • Valuation-model inputs (compare live rates vs your hardcoded assumptions)
"""

import json
from datetime import datetime
from html import escape
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
IN_FILE  = BASE_DIR / "data" / "macro.json"
OUT_FILE = BASE_DIR / "dashboards" / "macro_dashboard.html"


def fmt_pct(v, dp=2):
    if v is None:
        return "—"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.{dp}f}%"

def fmt_num(v, dp=2):
    if v is None:
        return "—"
    return f"{v:,.{dp}f}"

def stale_badge(date_str: str, threshold_days: int = 365) -> str:
    """Return a small yellow badge if the observation date is older than threshold."""
    if not date_str:
        return ""
    try:
        obs = datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        return ""
    age = (datetime.now() - obs).days
    if age > threshold_days:
        years = age / 365
        return f' <span class="stale">stale ({years:.1f}y old)</span>'
    return ""


def fmt_delta(curr, prev, dp=2, prefix=""):
    if curr is None or prev is None:
        return ""
    d = curr - prev
    sign = "+" if d >= 0 else ""
    cls = "pos" if d >= 0 else "neg"
    return f'<span class="delta {cls}">{sign}{d:.{dp}f}{prefix}</span>'

def cls_for(v):
    if v is None:
        return ""
    return "pos" if v >= 0 else "neg"


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
h2 { margin: 0 0 12px 0; font-size: 16px; color: #e2e8f0; }
.sub { color: #94a3b8; font-size: 13px; margin-bottom: 24px; }
.grid-cards {
    display: grid; gap: 12px;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    margin-bottom: 24px;
}
.card { background: #1e293b; border: 1px solid #334155; border-radius: 10px; padding: 14px; }
.card .label { color: #94a3b8; font-size: 11px; text-transform: uppercase; letter-spacing: .06em; }
.card .value { font-size: 20px; font-weight: 600; margin-top: 4px; }
.card .sub2  { font-size: 12px; color: #94a3b8; margin-top: 3px; }
.section {
    background: #1e293b; border: 1px solid #334155;
    border-radius: 10px; padding: 16px; margin-bottom: 16px;
}
.two-col {
    display: grid; gap: 16px;
    grid-template-columns: repeat(auto-fit, minmax(480px, 1fr));
    margin-bottom: 16px;
}
.two-col .section { margin-bottom: 0; }
.three-col {
    display: grid; gap: 16px;
    grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
    margin-bottom: 16px;
}
.three-col .section { margin-bottom: 0; }
.section table { margin-bottom: 0; }
.section h2 { margin-bottom: 8px; }
th, td { padding: 6px 8px; }
.section .meta { color: #94a3b8; font-size: 12px; margin-bottom: 10px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { padding: 7px 10px; text-align: left; border-bottom: 1px solid #334155; }
th { color: #94a3b8; font-weight: 500; font-size: 11px; text-transform: uppercase; letter-spacing: .05em; }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
tbody tr:hover { background: #283548; }
.pos { color: #10b981; }
.neg { color: #ef4444; }
.delta { font-size: 12px; margin-left: 6px; }
.region { display: inline-block; padding: 1px 8px; border-radius: 999px;
          background: #0f172a; border: 1px solid #334155;
          font-size: 10px; color: #94a3b8; text-transform: uppercase;
          letter-spacing: .05em; margin-right: 6px; }
.warn { color: #fbbf24; }
.stale {
    display: inline-block; padding: 1px 6px; border-radius: 4px;
    background: #78350f; color: #fbbf24;
    font-size: 10px; font-weight: 500; letter-spacing: 0.04em;
    margin-left: 4px;
}
.footer { color: #64748b; font-size: 11px; margin-top: 32px; text-align: center; }
"""


def headline_cards(data):
    """Pick out a handful of key series for the top strip."""
    fred    = data.get("fred", {})
    markets = data.get("markets", {})
    fx      = data.get("fx", {})

    cards = []

    def card(label, value, sub=""):
        cls = ""
        if sub and sub.startswith(("+", "-")):
            cls = " pos" if sub.startswith("+") else " neg"
        cards.append(f"""
          <div class="card">
            <div class="label">{escape(label)}</div>
            <div class="value">{escape(value)}</div>
            <div class="sub2{cls}">{escape(sub)}</div>
          </div>""")

    if (uk10 := fred.get("uk_10y")) and (val := uk10.get("value")) is not None:
        prev = uk10.get("prev")
        delta = f"{val - prev:+.2f} pp" if prev is not None else uk10.get("date", "")
        card("UK 10Y gilt", f"{val:.2f}%", delta)
    if (us10 := fred.get("us_10y")) and (val := us10.get("value")) is not None:
        prev = us10.get("prev")
        delta = f"{val - prev:+.2f} pp" if prev is not None else us10.get("date", "")
        card("US 10Y treasury", f"{val:.2f}%", delta)

    if (gbp := fx.get("gbpusd")) and (val := gbp.get("value")) is not None:
        card("GBP/USD", f"{val:.4f}", fmt_pct(gbp.get("change_pct")))
    if (dxy := fx.get("dxy")) and (val := dxy.get("value")) is not None:
        card("DXY", f"{val:.2f}", fmt_pct(dxy.get("change_pct")))

    if (ftse := markets.get("ftse100")) and (val := ftse.get("value")) is not None:
        card("FTSE 100", fmt_num(val, 0), fmt_pct(ftse.get("change_pct")))
    if (spx := markets.get("sp500")) and (val := spx.get("value")) is not None:
        card("S&P 500", fmt_num(val, 0), fmt_pct(spx.get("change_pct")))
    if (vix := markets.get("vix")) and (val := vix.get("value")) is not None:
        card("VIX", f"{val:.2f}", fmt_pct(vix.get("change_pct")))

    return '<div class="grid-cards">' + "".join(cards) + '</div>'


def events_table(events):
    if not events:
        return '<tr><td colspan="7" style="color:var(--ink-3);text-align:center;padding:12px">No upcoming releases (check FINNHUB_API_KEY).</td></tr>'

    from datetime import datetime as _dt
    now = _dt.now()

    def fmt_val(v, unit=""):
        if v in (None, ""):
            return "—"
        try:
            fv = float(v)
        except (TypeError, ValueError):
            return escape(str(v))
        if fv.is_integer() and abs(fv) >= 10:
            return f"{int(fv):,}{unit}"
        return f"{fv:.2f}{unit}".rstrip("0").rstrip(".") + ("" if (unit or "") in ("",) else "")

    def impact_badge(imp):
        imp = (imp or "").lower()
        if imp == "high":
            return '<span class="stale" style="background:#f5d0c7;color:#8a2f1a;border-color:#c77759">HIGH</span>'
        if imp == "medium":
            return '<span class="chip" style="background:#f5e6c9;color:#7c5a18">MED</span>'
        return f'<span class="chip" style="opacity:.6">{escape(imp.upper()) if imp else "—"}</span>'

    rows = []
    for e in events[:40]:
        ts = e.get("time", "")
        # Friendly "Mon 21 Apr · 12:30" format
        when = ts
        try:
            dt = _dt.strptime(ts[:16], "%Y-%m-%d %H:%M")
            when = dt.strftime("%a %d %b · %H:%M")
            if dt.date() == now.date():
                when = "Today · " + dt.strftime("%H:%M")
            elif (dt.date() - now.date()).days == 1:
                when = "Tomorrow · " + dt.strftime("%H:%M")
        except (ValueError, TypeError):
            pass

        actual = e.get("actual")
        estimate = e.get("estimate")
        prev = e.get("prev")
        cls_actual = ""
        if actual not in (None, "") and estimate not in (None, ""):
            try:
                cls_actual = "pos" if float(actual) > float(estimate) else ("neg" if float(actual) < float(estimate) else "")
            except (TypeError, ValueError):
                cls_actual = ""

        unit = e.get("unit") or ""
        rows.append(
            "<tr>"
            f'<td class="num" style="color:var(--ink-3)">{escape(when)}</td>'
            f'<td style="color:var(--ink-3)"><small>{escape(e.get("country",""))}</small></td>'
            f'<td>{escape(e.get("event",""))}</td>'
            f'<td class="num {cls_actual}">{fmt_val(actual, unit)}</td>'
            f'<td class="num" style="color:var(--ink-3)">{fmt_val(estimate, unit)}</td>'
            f'<td class="num" style="color:var(--ink-3)">{fmt_val(prev, unit)}</td>'
            f'<td>{impact_badge(e.get("impact"))}</td>'
            "</tr>"
        )
    return "".join(rows)


def rates_table(fred):
    rates = [
        ("us_fed_funds",   "Fed Funds"),
        ("uk_bank_rate",   "UK 3M"),
        ("ez_rate",        "ECB deposit"),
        ("us_10y",         "US 10Y"),
        ("uk_10y",         "UK 10Y"),
        ("de_10y",         "German 10Y"),
        ("jp_10y",         "Japan 10Y"),
    ]
    rows = []
    for key, short in rates:
        d = fred.get(key)
        if not d:
            continue
        val = d.get("value"); prev = d.get("prev")
        delta_html = fmt_delta(val, prev, dp=2, prefix=" pp")
        date = d.get("date", "")
        rows.append(f"""
          <tr>
            <td>{escape(d.get("label", short))}</td>
            <td class="num">{fmt_pct(val)}</td>
            <td class="num">{delta_html}</td>
            <td class="num">{escape(date)}{stale_badge(date)}</td>
          </tr>""")
    return "".join(rows) or '<tr><td colspan="4">No data (check FRED_API_KEY).</td></tr>'


def fx_table(fx):
    rows = []
    for key, d in fx.items():
        val = d.get("value"); change = d.get("change_pct")
        rows.append(f"""
          <tr>
            <td>{escape(d.get("label", key))}</td>
            <td class="num">{fmt_num(val, 4 if val and val < 10 else 2)}</td>
            <td class="num {cls_for(change)}">{fmt_pct(change)}</td>
            <td class="num">{escape(d.get("as_of", ""))}</td>
          </tr>""")
    return "".join(rows) or '<tr><td colspan="4">No FX data.</td></tr>'


def markets_table(markets):
    by_region = {}
    for key, d in markets.items():
        r = d.get("region", "Other")
        by_region.setdefault(r, []).append(d)
    rows = []
    for region in ["UK", "US", "Eurozone", "Germany", "France",
                   "Japan", "Hong Kong", "China", "Global"]:
        for d in by_region.get(region, []):
            val = d.get("value"); change = d.get("change_pct")
            rows.append(f"""
              <tr>
                <td><span class="region">{escape(region)}</span>{escape(d.get("label", ""))}</td>
                <td class="num">{fmt_num(val, 2)}</td>
                <td class="num {cls_for(change)}">{fmt_pct(change)}</td>
                <td class="num">{escape(d.get("as_of", ""))}</td>
              </tr>""")
    return "".join(rows) or '<tr><td colspan="4">No market data.</td></tr>'


def econ_table(fred):
    regions = [
        ("UK",       [("uk_gdp_yoy","GDP (YoY)"),
                      ("uk_cpi_yoy","CPI (YoY)"),
                      ("uk_unemployment","Unemployment")]),
        ("US",       [("us_gdp_yoy","GDP (YoY)"),
                      ("us_cpi_yoy","CPI (YoY)"),
                      ("us_unemployment","Unemployment")]),
        ("Eurozone", [("ez_gdp_yoy","GDP (YoY)"),
                      ("ez_cpi_yoy","HICP (YoY)"),
                      ("ez_unemployment","Unemployment")]),
        ("Japan",    [("jp_gdp_yoy","GDP (YoY)"),
                      ("jp_cpi_yoy","CPI (YoY)"),
                      ("jp_unemployment","Unemployment")]),
        ("China",    [("cn_gdp_yoy","GDP (YoY)"),
                      ("cn_cpi_yoy","CPI (YoY)")]),
    ]
    rows = []
    for region, metrics in regions:
        first = True
        for key, short in metrics:
            d = fred.get(key)
            if not d:
                continue
            val = d.get("value"); prev = d.get("prev")
            delta_html = fmt_delta(val, prev, dp=2, prefix=" pp")
            region_cell = (f'<span class="region">{escape(region)}</span>' if first else "")
            date = d.get("date", "")
            rows.append(f"""
              <tr>
                <td>{region_cell}{escape(short)}</td>
                <td class="num">{fmt_pct(val)}</td>
                <td class="num">{delta_html}</td>
                <td class="num">{escape(date)}{stale_badge(date)}</td>
              </tr>""")
            first = False
    return "".join(rows) or '<tr><td colspan="4">No economic data (check FRED_API_KEY).</td></tr>'


def model_table(data):
    """Compare live rates against monthly averages (FRED) and the hardcoded
    valuation-model assumptions. Shows three columns where data is available
    so it's clear when live has drifted from the smoother monthly read."""
    fred = data.get("fred", {})
    yl   = data.get("yields_live", {})
    a    = data.get("assumptions", {}) or {}

    def row(label, model_val, monthly_val, live_val, monthly_meta, live_meta):
        if model_val is None:
            return ""
        # Drift compares live (preferred) to model; falls back to monthly if no live
        ref = live_val if live_val is not None else monthly_val
        drift = (ref - model_val) if (ref is not None) else None
        drift_html = fmt_delta(ref, model_val, dp=2, prefix=" pp") if ref is not None else "—"
        status = ""
        if drift is not None and abs(drift) > 0.5:
            status = ' <span class="warn">Drifted &gt; 0.5 pp — consider updating valuation.py</span>'
        # As-of / source: prefer live's, else monthly's
        meta = live_meta or monthly_meta or ""
        return f"""
          <tr>
            <td>{escape(label)}</td>
            <td class="num">{fmt_pct(model_val)}</td>
            <td class="num">{fmt_pct(monthly_val) if monthly_val is not None else "—"}</td>
            <td class="num">{fmt_pct(live_val) if live_val is not None else "—"}</td>
            <td class="num">{drift_html}</td>
            <td>{escape(meta)}{status}</td>
          </tr>"""

    # UK 10y - monthly avg from FRED, live from BoE
    uk10_monthly = fred.get("uk_10y", {}).get("value")
    uk10_monthly_date = fred.get("uk_10y", {}).get("date", "")
    uk10_live_d = yl.get("uk_10y_live", {}) or {}
    uk10_live = uk10_live_d.get("value")
    uk10_live_meta = (f"Live: BoE IADB ({uk10_live_d.get('as_of', '?')})"
                      if uk10_live is not None else "")
    uk10_monthly_meta = f"Monthly avg: FRED ({uk10_monthly_date})" if uk10_monthly_date else ""

    # US 10y - monthly avg from FRED, live from yfinance
    us10_monthly = fred.get("us_10y", {}).get("value")
    us10_monthly_date = fred.get("us_10y", {}).get("date", "")
    us10_live_d = yl.get("us_10y_live", {}) or {}
    us10_live = us10_live_d.get("value")
    us10_live_meta = (f"Live: yfinance ^TNX ({us10_live_d.get('as_of', '?')})"
                      if us10_live is not None else "")
    us10_monthly_meta = f"Monthly avg: FRED ({us10_monthly_date})" if us10_monthly_date else ""

    rows = [
        row("Risk-free rate (UK)", a.get("rf_uk_pct"),
            uk10_monthly, uk10_live, uk10_monthly_meta, uk10_live_meta),
        row("Risk-free rate (US)", a.get("rf_us_pct"),
            us10_monthly, us10_live, us10_monthly_meta, us10_live_meta),
        row("Equity risk premium", a.get("erp_pct"),
            None, None, "", "Hard to observe — check Damodaran's ERP page"),
        row("Terminal growth",     a.get("terminal_g_pct"),
            None, 2.0, "", "BoE 2% inflation target"),
        row("5Y growth assumption", a.get("growth_5y_pct"),
            None, None, "", "Hardcoded in valuation.py"),
        row("WACC default",        a.get("wacc_default_pct"),
            None, None, "", "Hardcoded in valuation.py"),
    ]
    return "".join(rows)


def build_html(data):
    generated_at = data.get("generated_at", "")
    try:
        when = datetime.fromisoformat(generated_at).strftime("%d %b %Y %H:%M")
    except Exception:
        when = generated_at

    sections = f"""
    <div class="section">
      <h2>Headline</h2>
      {headline_cards(data)}
    </div>

    <div class="three-col">
      <div class="section">
        <h2>Rates &amp; yields</h2>
        <table>
          <thead><tr>
            <th>Series</th><th class="num">Level</th><th class="num">Δ</th><th class="num">As of</th>
          </tr></thead>
          <tbody>{rates_table(data.get("fred", {}))}</tbody>
        </table>
      </div>

      <div class="section">
        <h2>FX</h2>
        <table>
          <thead><tr>
            <th>Pair</th><th class="num">Rate</th><th class="num">Δ 1d</th><th class="num">As of</th>
          </tr></thead>
          <tbody>{fx_table(data.get("fx", {}))}</tbody>
        </table>
      </div>

      <div class="section">
        <h2>Equity markets</h2>
        <table>
          <thead><tr>
            <th>Index</th><th class="num">Level</th><th class="num">Δ 1d</th><th class="num">As of</th>
          </tr></thead>
          <tbody>{markets_table(data.get("markets", {}))}</tbody>
        </table>
      </div>
    </div>

    <div class="two-col">
      <div class="section">
        <h2>Economic indicators</h2>
        <div class="meta">GDP / CPI / Unemployment — latest FRED observations.</div>
        <table>
          <thead><tr>
            <th>Region / metric</th><th class="num">Latest</th><th class="num">Δ</th><th class="num">As of</th>
          </tr></thead>
          <tbody>{econ_table(data.get("fred", {}))}</tbody>
        </table>
      </div>

      <div class="section">
        <h2>Valuation-model inputs</h2>
        <div class="meta">
          Assumptions baked into <code>scripts/valuation.py</code>.
          Drift &gt;0.5 pp flags a potential update.
        </div>
        <table>
          <thead><tr>
            <th>Input</th><th class="num">Model</th><th class="num">Monthly avg</th><th class="num">Live</th><th class="num">Drift (live - model)</th><th>As of / source</th>
          </tr></thead>
          <tbody>{model_table(data)}</tbody>
        </table>
      </div>
    </div>

    <div class="section">
      <h2>Upcoming economic releases</h2>
      <div class="meta">
        GB / US / EU release calendar · next 21 days · from Finnhub
        (same feed powering investing.com). Actual vs. estimate vs. previous.
      </div>
      <table>
        <thead><tr>
          <th class="num" style="width:110px">When</th>
          <th style="width:50px">Cty</th>
          <th>Event</th>
          <th class="num">Actual</th>
          <th class="num">Estimate</th>
          <th class="num">Previous</th>
          <th>Impact</th>
        </tr></thead>
        <tbody>{events_table(data.get("events", []))}</tbody>
      </table>
    </div>
    """

    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Macro Dashboard</title>
{_FONTS_LINK}
<style>{_EDITORIAL_CSS}</style>
</head><body>
{_nav_html(active="macro")}
<h1>Macro Dashboard</h1>
<div class="sub">Global rates, FX, equities &amp; economic indicators · generated {escape(when)}</div>
{sections}
<div class="footer">
  Data: FRED (macro series) · Yahoo Finance (markets + FX) · Finnhub (release calendar).
  FRED series update monthly or quarterly; markets + FX are previous-close;
  release calendar refreshes with each hourly macro sync.
</div>
{_THEME_JS}
</body></html>"""


def main():
    if not IN_FILE.exists():
        print(f"  ERROR: {IN_FILE} not found — run sync_macro.py first")
        return

    data = json.loads(IN_FILE.read_text(encoding="utf-8"))
    OUT_FILE.write_text(build_html(data), encoding="utf-8")
    print(f"  Wrote {OUT_FILE}")


if __name__ == "__main__":
    main()
