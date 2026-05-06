"""generate_exposure_dashboard.py — Render the consolidated stock-exposure
dashboard (eToro + T212 direct holdings + fund/ETF decomposition + LBG pension
proxy).

Reads:  data/stock_exposure.json
Writes: dashboards/exposure_dashboard.html
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent.parent
DATA = BASE / "data"
OUT = BASE / "dashboards" / "exposure_dashboard.html"

sys.path.insert(0, str(Path(__file__).parent))
try:
    from editorial_theme import (
        CSS as _EDITORIAL_CSS,
        FONTS_LINK as _FONTS_LINK,
        nav_html as _nav_html,
        THEME_JS as _THEME_JS,
    )
except ImportError:
    from scripts.editorial_theme import (
        CSS as _EDITORIAL_CSS,
        FONTS_LINK as _FONTS_LINK,
        nav_html as _nav_html,
        THEME_JS as _THEME_JS,
    )


def fmt_gbp(v: float | None, dp: int = 0) -> str:
    if v is None:
        return "—"
    sign = "-" if v < 0 else ""
    return f"{sign}£{abs(v):,.{dp}f}"


def _escape(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


EXTRA_CSS = """
.exp-hero { display: grid; gap: 10px; margin-bottom: 18px;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); }
.exp-hero .kpi { padding: 14px 16px; }
.kpi-sub { font-family: var(--f-mono); font-size: 10px; color: var(--ink-3);
           margin-top: 6px; text-transform: uppercase; letter-spacing: 0.08em; }
.exp-row { display: grid; gap: 16px; margin-bottom: 16px;
           grid-template-columns: 1.5fr 1fr; }
@media (max-width: 1000px) { .exp-row { grid-template-columns: 1fr; } }
.bar-fill { height: 100%; background: var(--accent-2); border-radius: 2px; }
.bar-track { background: var(--bg-sunken); height: 6px; border-radius: 2px;
             overflow: hidden; margin-top: 3px; }
.exp-table td.bar-cell { width: 140px; padding: 4px 10px !important; }
.exp-table .direct { color: var(--ink-2); }
.exp-table .indirect { color: var(--ink-3); font-style: italic; }
.exp-rest { color: var(--ink-3); font-size: 11px; font-style: italic; }
.chip-via { font-family: var(--f-mono); font-size: 9px; padding: 1px 6px;
            background: var(--bg-sunken); border: 1px solid var(--line);
            border-radius: 10px; color: var(--ink-3); letter-spacing: .06em;
            text-transform: uppercase; margin-left: 4px; }

details.fund { margin-bottom: 10px; background: var(--bg-card);
               border: 1px solid var(--line); border-radius: 4px; padding: 10px 14px; }
details.fund > summary { cursor: pointer; list-style: none; }
details.fund > summary::-webkit-details-marker { display: none; }
details.fund > summary::before { content: "▸  "; color: var(--accent); font-family: var(--f-mono); }
details.fund[open] > summary::before { content: "▾  "; }
details.fund .fund-head {
    display: flex; justify-content: space-between; align-items: baseline; gap: 12px;
    font-family: var(--f-display); font-size: 16px;
}
details.fund .fund-coverage {
    font-family: var(--f-mono); font-size: 10px; color: var(--ink-3);
    text-transform: uppercase; letter-spacing: .08em;
}
details.fund table { margin-top: 10px; font-size: 12px; }
details.fund td { padding: 4px 10px !important; }

.source-chip {
    display: inline-block; padding: 2px 8px; margin: 0 4px 3px 0;
    background: var(--bg-sunken); border: 1px solid var(--line);
    border-radius: 12px; font-family: var(--f-mono); font-size: 10px;
    color: var(--ink-3); letter-spacing: 0.05em;
}
.filter-row {
    display: flex; gap: 10px; align-items: center; margin: 8px 0 12px;
    font-family: var(--f-mono); font-size: 11px; color: var(--ink-3);
    text-transform: uppercase; letter-spacing: .08em;
}
.filter-row input[type="checkbox"] { accent-color: var(--accent); }
.filter-row label { cursor: pointer; }
.search-inp {
    font-family: var(--f-mono); font-size: 12px; padding: 4px 8px;
    background: var(--bg-card); border: 1px solid var(--line); color: var(--ink);
    border-radius: 3px; min-width: 180px;
}
"""


def render_kpi_strip(totals: dict) -> str:
    def _kpi(label, value, sub=""):
        return (
            '<div class="kpi">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            f'<div class="kpi-sub">{sub}</div>'
            '</div>'
        )
    equity = totals.get("equity_gbp", 0)
    direct = totals.get("direct_gbp", 0)
    indirect = totals.get("indirect_gbp", 0)
    grand = totals.get("grand_total_gbp", 0) or 1
    return (
        '<div class="exp-hero">'
        + _kpi("Total Equity Exposure", fmt_gbp(equity),
               f"{equity / grand * 100:.1f}% of grand total")
        + _kpi("Direct Single Stocks", fmt_gbp(direct),
               f"{direct / equity * 100:.1f}% of equity" if equity else "")
        + _kpi("Indirect (via funds)", fmt_gbp(indirect),
               f"{indirect / equity * 100:.1f}% of equity" if equity else "")
        + _kpi("Bonds + Cash + Crypto", fmt_gbp(
            totals.get("bonds_gbp", 0) + totals.get("cash_gbp", 0)
            + totals.get("crypto_gbp", 0) + totals.get("alt_gbp", 0)),
            "non-equity")
        + '</div>'
    )


def render_consolidated(rows: list[dict], total_equity: float) -> str:
    """Main table — one row per ticker, sorted by total exposure."""
    max_val = max((r["total_gbp"] for r in rows), default=1)

    def _row(r: dict, idx: int) -> str:
        is_rest = r["ticker"].startswith("__REST_")
        ticker_display = "— residual —" if is_rest else r["ticker"]
        name_cls = 'exp-rest' if is_rest else ''
        bar_pct = (r["total_gbp"] / max_val * 100) if max_val else 0
        # Source chips
        src_counts: dict[str, float] = {}
        for s in r.get("sources", []):
            key = s.get("source", "?")
            src_counts[key] = src_counts.get(key, 0) + (s.get("value_gbp") or 0)
        chips = "".join(
            f'<span class="source-chip">{_escape(k)} {fmt_gbp(v)}</span>'
            for k, v in sorted(src_counts.items(), key=lambda x: -x[1])
        )
        return f"""
        <tr data-rest="{1 if is_rest else 0}"
            data-ticker="{_escape(r['ticker'].lower())}"
            data-name="{_escape((r['name'] or '').lower())}">
          <td class="num" style="color:var(--ink-3);font-size:11px">{idx}</td>
          <td><code style="font-family:var(--f-mono);font-size:12px">{_escape(ticker_display)}</code></td>
          <td class="{name_cls}">{_escape(r['name'])}<div style="margin-top:3px">{chips}</div></td>
          <td class="num direct">{fmt_gbp(r['direct_gbp']) if r['direct_gbp'] else '—'}</td>
          <td class="num indirect">{fmt_gbp(r['indirect_gbp']) if r['indirect_gbp'] else '—'}</td>
          <td class="num" style="font-weight:500">{fmt_gbp(r['total_gbp'])}</td>
          <td class="num">{r['pct_of_equity']:.2f}%</td>
          <td class="bar-cell">
            <div class="bar-track"><div class="bar-fill" style="width:{bar_pct:.1f}%"></div></div>
          </td>
        </tr>"""

    body = "".join(_row(r, i + 1) for i, r in enumerate(rows))
    return f"""
    <div class="section">
      <h2>Consolidated Exposure</h2>
      <div class="filter-row">
        <label><input type="checkbox" id="hide-rest" checked> Hide fund residuals</label>
        <label><input type="checkbox" id="direct-only"> Direct only</label>
        <input class="search-inp" id="exp-search" placeholder="Filter by ticker or name…">
        <span style="flex:1"></span>
        <span id="row-count">{len(rows)} tickers</span>
      </div>
      <table class="exp-table">
        <thead>
          <tr>
            <th>#</th>
            <th>Ticker</th>
            <th>Name · sources</th>
            <th class="num">Direct</th>
            <th class="num">Indirect</th>
            <th class="num">Total</th>
            <th class="num">% of equity</th>
            <th></th>
          </tr>
        </thead>
        <tbody id="exp-rows">{body}</tbody>
      </table>
    </div>
    """


def render_funds(funds: list[dict]) -> str:
    """Collapsible list of each fund with its top-10 holdings."""
    if not funds:
        return ""
    # Group by ticker — a fund may appear in multiple accounts.
    grouped: dict[str, dict] = {}
    for f in funds:
        k = f["ticker"]
        if k not in grouped:
            grouped[k] = {
                "ticker": k, "name": f["name"],
                "total_value_gbp": 0.0, "sources": [],
                "top_holdings": f.get("top_holdings", []),
                "covered_pct": f.get("covered_pct", 0.0),
            }
        grouped[k]["total_value_gbp"] += f.get("total_value_gbp") or 0
        grouped[k]["sources"].append({
            "source": f.get("source", ""),
            "value_gbp": f.get("total_value_gbp") or 0,
        })

    items = sorted(grouped.values(), key=lambda x: x["total_value_gbp"], reverse=True)

    def _fund_html(f: dict) -> str:
        sources_html = " ".join(
            f'<span class="source-chip">{_escape(s["source"])} {fmt_gbp(s["value_gbp"])}</span>'
            for s in f["sources"]
        )
        hold_rows = "".join(
            f'<tr><td><code style="font-family:var(--f-mono);font-size:11px">{_escape(h["ticker"])}</code></td>'
            f'<td>{_escape(h["name"])}</td>'
            f'<td class="num">{h["pct"]*100:.2f}%</td>'
            f'<td class="num">{fmt_gbp(f["total_value_gbp"] * h["pct"])}</td></tr>'
            for h in f["top_holdings"]
        )
        residual = 1 - f.get("covered_pct", 0)
        if residual > 0.001:
            hold_rows += (
                f'<tr><td colspan="2" class="exp-rest">(Rest of fund — beyond top-10)</td>'
                f'<td class="num">{residual*100:.2f}%</td>'
                f'<td class="num">{fmt_gbp(f["total_value_gbp"] * residual)}</td></tr>'
            )
        return f"""
        <details class="fund">
          <summary>
            <div class="fund-head">
              <div>
                <code style="font-family:var(--f-mono);font-size:13px;color:var(--accent)">{_escape(f["ticker"])}</code>
                <span style="margin-left:8px">{_escape(f["name"])}</span>
              </div>
              <div>
                <span class="fund-coverage">{f.get("covered_pct", 0)*100:.0f}% top-10 coverage</span>
                <span style="font-family:var(--f-mono);margin-left:12px">{fmt_gbp(f["total_value_gbp"])}</span>
              </div>
            </div>
            <div style="margin-top:4px">{sources_html}</div>
          </summary>
          <table>
            <thead>
              <tr><th>Ticker</th><th>Holding</th><th class="num">Weight</th><th class="num">Value</th></tr>
            </thead>
            <tbody>{hold_rows}</tbody>
          </table>
        </details>
        """

    items_html = "".join(_fund_html(f) for f in items)
    return f"""
    <div class="section">
      <h2>Funds &amp; ETFs — Top-10 Holdings</h2>
      <small style="color:var(--ink-3);font-family:var(--f-mono);font-size:11px;
                    text-transform:uppercase;letter-spacing:0.08em">
        Decomposed via yfinance. LBG Global Equity Fund proxied as average of IWDA.L + VWRL.L.
      </small>
      <div style="margin-top:14px">{items_html}</div>
    </div>
    """


def render_unresolved(rows: list[dict]) -> str:
    if not rows:
        return ""
    body = "".join(
        f'<tr><td>{_escape(r["source"])}</td>'
        f'<td>{_escape(r["name"])}</td>'
        f'<td><code>{_escape(r.get("ticker",""))}</code></td>'
        f'<td class="num">{fmt_gbp(r["value_gbp"])}</td>'
        f'<td style="color:var(--accent);font-family:var(--f-mono);font-size:10px">{_escape(r["reason"])}</td></tr>'
        for r in rows
    )
    return f"""
    <div class="section" style="margin-top:16px">
      <h2>Unresolved Positions</h2>
      <small style="color:var(--ink-3)">These positions could not be automatically decomposed
        — value counted as direct but may need a manual ticker mapping.</small>
      <table>
        <thead><tr><th>Source</th><th>Name</th><th>Ticker</th><th class="num">Value</th><th>Reason</th></tr></thead>
        <tbody>{body}</tbody>
      </table>
    </div>
    """


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    if not (DATA / "stock_exposure.json").exists():
        print("  WARN: data/stock_exposure.json not found — run sync_stock_exposure.py first")
        return
    d = json.loads((DATA / "stock_exposure.json").read_text(encoding="utf-8"))
    totals = d.get("totals") or {}
    consolidated = d.get("consolidated") or []
    funds = d.get("funds") or []
    unresolved = d.get("unresolved") or []
    generated = d.get("generated_at", datetime.now().isoformat(timespec="seconds"))

    # Filter sensible: hide zero-value rows
    rows = [r for r in consolidated if (r.get("total_gbp") or 0) > 0.5]

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Stock Exposure Dashboard</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
{_FONTS_LINK}
<style>{_EDITORIAL_CSS}{EXTRA_CSS}</style>
</head>
<body>
{_nav_html(active="exposure", privacy=True)}
<main>
  <h1>Stock Exposure</h1>
  <div class="sub">Consolidated across eToro, Trading 212, AJ Bell, LBG Pension · Generated {generated}</div>

  {render_kpi_strip(totals)}

  {render_consolidated(rows, totals.get("equity_gbp", 1))}

  {render_funds(funds)}

  {render_unresolved(unresolved)}

  <div class="footer">
    Data sources: combined_portfolio.json (eToro + T212) + finances.json (Finances ND sheet).
    Fund decomposition via yfinance top-10 holdings. LBG Global Equity Fund proxied as
    50/50 IWDA.L + VWRL.L. Residual = sheet value not covered by top-10.
  </div>
</main>

{_THEME_JS}

<script>
(function(){{
  var hideRest = document.getElementById('hide-rest');
  var directOnly = document.getElementById('direct-only');
  var search = document.getElementById('exp-search');
  var rows = document.querySelectorAll('#exp-rows tr');
  var countSpan = document.getElementById('row-count');

  function apply(){{
    var hr = hideRest && hideRest.checked;
    var doOnly = directOnly && directOnly.checked;
    var q = (search && search.value || '').trim().toLowerCase();
    var visible = 0;
    rows.forEach(function(r){{
      var show = true;
      if (hr && r.getAttribute('data-rest') === '1') show = false;
      if (doOnly) {{
        var direct = r.cells[3].textContent.trim();
        if (!direct || direct === '—') show = false;
      }}
      if (q) {{
        var tk = r.getAttribute('data-ticker') || '';
        var nm = r.getAttribute('data-name') || '';
        if (tk.indexOf(q) === -1 && nm.indexOf(q) === -1) show = false;
      }}
      r.style.display = show ? '' : 'none';
      if (show) visible++;
    }});
    if (countSpan) countSpan.textContent = visible + ' tickers';
  }}

  if (hideRest) hideRest.addEventListener('change', apply);
  if (directOnly) directOnly.addEventListener('change', apply);
  if (search) search.addEventListener('input', apply);
  apply();
}})();
</script>
</body></html>
"""
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(html, encoding="utf-8")
    print(f"  Wrote {OUT}")


if __name__ == "__main__":
    main()
