"""generate_finances_dashboard.py — Render a personal-wealth dashboard from
data/finances.json (produced by sync_finances.py)."""
from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from datetime import datetime

try:
    from editorial_theme import CSS as _EDITORIAL_CSS, FONTS_LINK as _FONTS_LINK, nav_html as _nav_html, THEME_JS as _THEME_JS
except ImportError:
    from scripts.editorial_theme import CSS as _EDITORIAL_CSS, FONTS_LINK as _FONTS_LINK, nav_html as _nav_html, THEME_JS as _THEME_JS

BASE_DIR = Path(__file__).parent.parent
DATA = BASE_DIR / "data" / "finances.json"
OUTPUT = BASE_DIR / "dashboards" / "finances_dashboard.html"


# ── Formatting helpers ────────────────────────────────────────────────────────
def fmt_gbp(v, dp: int = 0) -> str:
    if v is None:
        return "—"
    sign = "-" if v < 0 else ""
    return f"{sign}£{abs(v):,.{dp}f}"


def fmt_pct(v, dp: int = 1) -> str:
    if v is None:
        return "—"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.{dp}f}%"


def colour_cls(v) -> str:
    if v is None:
        return ""
    return "pos" if v >= 0 else "neg"


_MONEY_RE = __import__("re").compile(r"[£$,\s]")

def parse_money_val(v):
    """Parse '£1,234' / '-£12' → float. For use inside the generator when
    rolling up values from the raw Portfolio grid."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in {"-", "—", "N/A"}:
        return None
    neg = s.startswith("-") or s.startswith("(-")
    s = _MONEY_RE.sub("", s.lstrip("-(").rstrip(")"))
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


# ── Section renderers ────────────────────────────────────────────────────────
def render_headline(nw: dict, investments: dict | None = None) -> str:
    cur = nw.get("current") or {}
    history = nw.get("history") or []

    total_assets = cur.get("assets")
    liab         = cur.get("liabilities")
    net          = cur.get("net")

    # Find the previous year snapshot (before "Current").
    prev = None
    for h in reversed(history):
        if (h.get("year") or "").lower() != "current" and h is not cur:
            prev = h
            break

    def _yoy(current, prior):
        if current is None or prior is None:
            return None, None
        d = current - prior
        pct = (d / prior * 100) if prior else None
        return d, pct

    assets_d, assets_p = _yoy(total_assets, (prev or {}).get("assets"))
    liab_d,   liab_p   = _yoy(liab,         (prev or {}).get("liabilities"))
    net_d,    net_p    = _yoy(net,          (prev or {}).get("net"))

    # Portfolio value + YoY from investments series
    inv_series = (investments or {}).get("series") or {}
    total_series = inv_series.get("Total") or []
    change_series = inv_series.get("Change") or []
    port_val = total_series[-1] if total_series else None
    port_d   = change_series[-1] if change_series else None
    port_prev = (port_val - port_d) if (port_val is not None and port_d is not None) else None
    port_p   = (port_d / port_prev * 100) if (port_d is not None and port_prev) else None

    def _card(label: str, value, val_colour, d, p, *, invert=False):
        """Render a single KPI card. invert=True means a DECREASE is good (liabilities)."""
        sub_cls = ""
        if d is not None:
            is_improvement = (d < 0) if invert else (d >= 0)
            sub_cls = "pos" if is_improvement else "neg"
        delta_sign = "+" if (d is not None and d >= 0) else ""
        delta_txt = f"{delta_sign}{fmt_gbp(d)}" if d is not None else "—"
        pct_txt = f"({delta_sign}{p:.1f}%)" if p is not None else ""
        return f"""
  <div class="kpi">
    <div class="kpi-label">{label}</div>
    <div class="kpi-value" style="{val_colour}">{fmt_gbp(value)}</div>
    <div class="kpi-sub {sub_cls}" style="margin-top:6px;font-family:var(--f-mono);font-size:12px;">
      {delta_txt} <span style="opacity:0.75;">{pct_txt} YoY</span>
    </div>
  </div>"""

    return f"""
<div class="kpi-grid">
  {_card('Total Assets',      total_assets, '', assets_d, assets_p)}
  {_card('Total Liabilities', liab,         'color:var(--accent)', liab_d, liab_p, invert=True)}
  {_card('Net Assets',        net,          'color:var(--accent-2);font-weight:600', net_d, net_p)}
  {_card('Portfolio Value',   port_val,     '', port_d, port_p)}
</div>
"""


def render_yoy_chart(nw: dict) -> str:
    hist = nw.get("history") or []
    rows = []
    for h in hist:
        yr = str(h.get("year"))
        assets = h.get("assets")
        liab = h.get("liabilities")
        net = h.get("net")
        chg = h.get("change")
        pct = (chg / (net - chg) * 100) if chg and net is not None and (net - chg) else None
        rows.append(
            f'<tr>'
            f'<td>{yr}</td>'
            f'<td class="num">{fmt_gbp(assets)}</td>'
            f'<td class="num" style="color:var(--text-2)">{fmt_gbp(liab)}</td>'
            f'<td class="num"><strong>{fmt_gbp(net)}</strong></td>'
            f'<td class="num {colour_cls(chg)}">{fmt_gbp(chg) if chg else "—"}</td>'
            f'<td class="num {colour_cls(pct)}">{fmt_pct(pct)}</td>'
            f'</tr>'
        )
    return f"""
<div class="card">
  <h2>Net Worth — Year on Year</h2>
  <table>
    <thead>
      <tr><th>Year</th><th class="num">Assets</th><th class="num">Liabilities</th><th class="num">Net</th><th class="num">YoY Change</th><th class="num">YoY %</th></tr>
    </thead>
    <tbody>{''.join(rows)}</tbody>
  </table>
</div>
"""


def render_allocation(portfolio: dict, retirement_wealth: dict, nw_current: dict | None = None) -> str:
    """Roll up holdings into 3 buckets: Pension, Investments, House Equity.
    The total equals Total Assets (net_worth.current.assets).
    """
    rows = (portfolio or {}).get("rows") or []

    def _cell(row, i):
        if not row or i >= len(row):
            return ""
        c = row[i]
        return str(c).strip() if c is not None else ""

    # Pass 1 — map each Investment-header row to the section name.
    # The section name lives in col_a of the header row OR in one of the next few rows.
    # Once we encounter a "Summary" / "Overall" block we stop (those are totals, not data).
    # NOTE: "Total Pensions." / "Total Other Investments" appear mid-document as per-section
    # sum rows and must NOT trigger the stop — only "Summary" / "Overall" do.
    section_for_row: dict[int, str] = {}
    current_section: str | None = None
    stopped = False
    for i, row in enumerate(rows):
        if stopped:
            continue
        col_a = _cell(row, 0)
        col_b = _cell(row, 1)
        a_low = col_a.lower()
        # Stop only at the Summary / Overall blocks at the bottom of the sheet
        if a_low in ("summary", "overall"):
            stopped = True
            continue
        # Skip (but don't stop) the per-section "Total X" marker rows
        if a_low.startswith("total "):
            continue
        if col_b.lower() == "investment":
            # New section header. Find its name: this row's col_a, or look forward up to 6 rows.
            section_name = col_a
            if not section_name:
                for j in range(i + 1, min(i + 7, len(rows))):
                    cand = _cell(rows[j], 0)
                    if cand and cand.lower() not in ("summary", "overall") and not cand.lower().startswith("total "):
                        section_name = cand
                        break
            current_section = section_name or current_section
            continue
        if current_section:
            section_for_row[i] = current_section

    # Pass 2 — sum holdings into buckets based on the section they belong to.
    pension_total = 0.0
    investments_total = 0.0
    for i, row in enumerate(rows):
        section = section_for_row.get(i)
        if not section:
            continue
        col_a = _cell(row, 0)
        col_b = _cell(row, 1)
        # Skip rows that are themselves section-name markers (col_a populated, col_b is the
        # first asset — keep these; but reject rows that are "Total X" accumulators).
        if col_b.lower().startswith("total "):
            continue
        if not col_b:
            continue
        # Skip stray section-name-only rows where col_a repeats a section label.
        if col_a.lower() in ("summary", "overall") or col_a.lower().startswith("total "):
            continue
        v = parse_money_val(row[6] if len(row) > 6 else None)
        if v is None:
            continue
        sec = section.lower()
        # Pension bucket = only LBG main pension ("Pensions" section) + AJ Bell Neil SIPP.
        # Everything else (ISAs, Dealing, LBG Employee shares, Alternative, Cash) is
        # non-pension Investments.
        is_pension = (
            sec == "pensions"
            or (sec.startswith("aj bell neil") and "isa" not in sec)
        )
        if is_pension:
            pension_total += v
        else:
            investments_total += v

    # House equity = Total Assets − Portfolio total (Pension + Investments).
    # This makes the allocation sum equal Total Assets exactly.
    nw_cur = nw_current if isinstance(nw_current, dict) else {}
    total_assets = nw_cur.get("assets")
    house_equity = None
    if total_assets is not None:
        house_equity = max(total_assets - pension_total - investments_total, 0)

    # Build buckets list, preserving order Pension → Investments → House Equity.
    items = [("Pension", pension_total), ("Investments", investments_total)]
    if house_equity is not None:
        items.append(("House Equity", house_equity))
    total = total_assets if total_assets is not None else sum(v for _, v in items)
    total = total or 1

    COLOURS = {"Pension": "#4f46e5", "Investments": "#10b981", "House Equity": "#f59e0b"}
    NOTES = {
        "Pension":       "LBG + AJ Bell Neil SIPP",
        "Investments":   "ISAs + Dealing + Alternative + Cash",
        "House Equity":  "Total Assets − Portfolio (estimated)",
    }
    row_html = "\n".join(
        f'<tr>'
        f'<td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{COLOURS.get(k, "#6366f1")};margin-right:7px;"></span>{k}'
        f'<br><small style="color:var(--muted)">{NOTES.get(k, "")}</small></td>'
        f'<td class="num">{fmt_gbp(v)}</td>'
        f'<td class="num">{v/total*100:.1f}%</td>'
        f'</tr>'
        for k, v in items
    )

    # Stacked bar visualising the split
    stacked = "".join(
        f'<div class="seg" style="width:{v/total*100:.1f}%;background:{COLOURS.get(k, "#6366f1")};">{k} {v/total*100:.0f}%</div>'
        for k, v in items if v > 0
    )

    return f"""
<div class="card">
  <h2>Asset Allocation</h2>
  <small style="color:var(--muted)">Three-way split matching Total Assets (£{total_assets:,.0f}).</small>
  <div class="stacked-bar" style="margin:14px 0 12px;">{stacked}</div>
  <table>
    <thead><tr><th>Bucket</th><th class="num">Value</th><th class="num">%</th></tr></thead>
    <tbody>{row_html}
      <tr style="border-top:2px solid var(--border);font-weight:600">
        <td>Total Assets</td><td class="num">{fmt_gbp(total)}</td><td class="num">100.0%</td>
      </tr>
    </tbody>
  </table>
</div>
"""


def render_retirement(retirement: dict, nw_current: float | None) -> str:
    scenarios = retirement.get("scenarios") or []
    projections = retirement.get("projections") or []

    # ── Table 1: Readiness per spend scenario ──
    rows_s = []
    for s in scenarios:
        gap = s.get("gap")
        gap_house = s.get("gap_with_house")
        months = s.get("months_to_save")
        # Colour: gap <= 0 = already there (pos); gap > 0 = not yet (neg)
        gap_cls = "pos" if (gap is not None and gap <= 0) else "neg"
        gap_h_cls = "pos" if (gap_house is not None and gap_house <= 0) else "neg"
        months_cls = "pos" if (months is not None and months <= 0) else ""
        if months is not None and months <= 0:
            months_display = "Already there"
        elif months is not None:
            months_display = f"{months:.1f} months ({months/12:.1f} yrs)"
        else:
            months_display = "—"
        rows_s.append(f"""
        <tr>
          <td><strong>{fmt_gbp(s['monthly'])}/mo</strong><br>
              <span style="color:var(--text-2);font-size:12px">{fmt_gbp(s['yearly'])}/yr at {s['spend_rate_pct']:.1f}% SWR</span></td>
          <td class="num">{fmt_gbp(s['pot_required'])}</td>
          <td class="num">{fmt_gbp(s['current_investments'])}</td>
          <td class="num {gap_cls}">{fmt_gbp(gap)}</td>
          <td class="num {gap_h_cls}">{fmt_gbp(gap_house)}</td>
          <td class="num {months_cls}">{months_display}</td>
        </tr>""")

    # ── Table 2: Pot projections at retirement ages ──
    rows_p = []
    for p in projections:
        rates = p.get("return_rates") or []
        continue_vals = p.get("pot_continue_invest") or []
        stop_vals = p.get("pot_stop_invest") or []
        months_remaining = p.get("months_remaining")
        years_remaining = (months_remaining / 12) if months_remaining else None
        rows_p.append(f"""
        <tr>
          <td rowspan="2" style="vertical-align:top">
            <strong>{p['label']}</strong><br>
            <span style="color:var(--text-2);font-size:12px">
              Retire {p.get('retirement_date','—')}<br>
              {fmt_gbp(p.get('current_pot'))} today<br>
              {(f'{years_remaining:.1f} yrs to go') if years_remaining else '—'}
            </span>
          </td>
          <td style="color:var(--text-2)">Continue investing</td>
          {''.join(f'<td class="num"><strong>{fmt_gbp(v)}</strong></td>' for v in continue_vals)}
        </tr>
        <tr>
          <td style="color:var(--text-2)">Stop investing today</td>
          {''.join(f'<td class="num">{fmt_gbp(v)}</td>' for v in stop_vals)}
        </tr>""")

    # Return-rate header comes from first projection's rates (they're all 5/8/10)
    rate_headers = ""
    if projections and projections[0].get("return_rates"):
        rate_headers = "".join(f'<th class="num">@ {r or "—"}</th>' for r in projections[0]["return_rates"])

    return f"""
<div class="card">
  <h2>Retirement Readiness</h2>
  <small style="color:var(--muted)">Three spending scenarios compared to current pot of {fmt_gbp(scenarios[0]['current_investments']) if scenarios else '—'}.</small>

  <h3 style="font-size:14px;margin:14px 0 6px 0;color:var(--text-2)">Spend Scenarios — Target vs Today</h3>
  <table>
    <thead>
      <tr>
        <th>Scenario</th>
        <th class="num">Pot Required</th>
        <th class="num">Current Investments</th>
        <th class="num">Gap (no house)</th>
        <th class="num">Gap (with house equity)</th>
        <th class="num">Months to Save</th>
      </tr>
    </thead>
    <tbody>{''.join(rows_s) if rows_s else '<tr><td colspan=6 style="color:var(--muted);text-align:center;padding:12px">No scenarios</td></tr>'}</tbody>
  </table>

  <h3 style="font-size:14px;margin:18px 0 6px 0;color:var(--text-2)">Projected Pot at Retirement — 55 / 57 / 60</h3>
  <table>
    <thead>
      <tr>
        <th>Scenario</th>
        <th>Strategy</th>
        {rate_headers}
      </tr>
    </thead>
    <tbody>{''.join(rows_p) if rows_p else '<tr><td colspan=5 style="color:var(--muted);text-align:center;padding:12px">No projections</td></tr>'}</tbody>
  </table>
</div>
"""


def render_school_fees(sf: dict) -> str:
    if not sf:
        return ""
    total = sf.get("total_cost") or 0
    paid = sf.get("paid") or 0
    saved = sf.get("saved") or 0
    remaining = sf.get("remaining") or 0
    predicted = sf.get("predicted_savings") or 0
    years_left = sf.get("years_remaining") or 0
    new_remaining = sf.get("new_remaining") or 0

    paid_pct = (paid / total * 100) if total else 0
    saved_pct = (saved / total * 100) if total else 0
    rem_pct = (remaining / total * 100) if total else 0

    paid_ytd = sf.get("paid_ytd") or 0
    ytd_months = sf.get("ytd_months") or 0
    ytd_year = sf.get("ytd_year") or ""
    ytd_label = f"Paid YTD ({ytd_year})" if ytd_year else "Paid YTD"
    ytd_meta = f"{ytd_months} month{'s' if ytd_months != 1 else ''}" if ytd_months else "—"

    # Net position after predicted future savings.
    # Positive surplus = predicted savings cover (and exceed) the still-to-fund amount.
    # The sheet stores `new_remaining` as absolute magnitude, so derive the sign here.
    surplus = predicted - remaining
    if surplus >= 0:
        net_label = "Surplus after predicted savings"
        net_cls = "pos"
    else:
        net_label = "Net shortfall after predicted savings"
        net_cls = "neg"
    net_value = abs(surplus)

    remaining_to_pay = max(total - paid, 0)
    bar_pct = min(paid_pct, 100)

    return f"""
<div class="card">
  <h2>School Fees Tracker</h2>
  <small style="color:var(--muted)">{years_left:.0f} years remaining • Total plan {fmt_gbp(total)}</small>

  <!-- Paid vs Expected progress bar -->
  <div style="margin:14px 0 18px;">
    <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px;font-size:12px;">
      <span><strong>{fmt_gbp(paid)}</strong> paid <span style="color:var(--muted)">of</span> <strong>{fmt_gbp(total)}</strong> expected</span>
      <span style="color:var(--muted)">{paid_pct:.1f}% • {fmt_gbp(remaining_to_pay)} remaining</span>
    </div>
    <div style="position:relative;height:10px;background:var(--bg-sunken, #ece8dd);border-radius:5px;overflow:hidden;border:1px solid var(--border, #dcd6c8);">
      <div style="position:absolute;inset:0;width:{bar_pct:.1f}%;background:linear-gradient(90deg,#8b5cf6,#6366f1);border-radius:5px;"></div>
    </div>
  </div>

  <table>
    <thead><tr><th>Component</th><th class="num">Amount</th><th class="num">% of plan</th></tr></thead>
    <tbody>
      <tr><td>Paid so far</td>                    <td class="num">{fmt_gbp(paid)}</td>     <td class="num">{paid_pct:.1f}%</td></tr>
      <tr><td style="color:var(--muted)">&nbsp;&nbsp;↳ {ytd_label}</td><td class="num">{fmt_gbp(paid_ytd)}</td><td class="num" style="color:var(--muted)">{ytd_meta}</td></tr>
      <tr><td>Already saved (not yet paid)</td>   <td class="num">{fmt_gbp(saved)}</td>    <td class="num">{saved_pct:.1f}%</td></tr>
      <tr><td>Still to fund</td>                  <td class="num">{fmt_gbp(remaining)}</td><td class="num">{rem_pct:.1f}%</td></tr>
      <tr style="border-top:2px solid var(--border);font-weight:600"><td>Total plan</td><td class="num">{fmt_gbp(total)}</td><td class="num">100.0%</td></tr>
      <tr><td style="color:var(--text-2)">Predicted future savings</td><td class="num">{fmt_gbp(predicted)}</td><td class="num">—</td></tr>
      <tr style="font-weight:600"><td>{net_label}</td><td class="num {net_cls}">{fmt_gbp(net_value)}</td><td class="num">—</td></tr>
    </tbody>
  </table>
</div>
"""


def extract_overall(portfolio: dict) -> dict:
    """Extract the 'Overall' grand-total row from the Portfolio tab.
    Returns {value, invested, inv_change, total_change, cap_roi, dividends, div_yield}.
    """
    rows = (portfolio or {}).get("rows") or []
    for r in rows:
        col_a = (r[0] if len(r) > 0 else "").strip().lower() if r else ""
        if col_a != "overall":
            continue
        def _at(idx):
            return r[idx] if idx < len(r) else None
        return {
            "value":        parse_money_val(_at(6)),
            "invested":     parse_money_val(_at(7)),
            "inv_change":   parse_money_val(_at(8)),
            "total_change": parse_money_val(_at(9)),
            "cap_roi":      _parse_pct_val(_at(10)),
            "dividends":    parse_money_val(_at(11)),
            "div_yield":    _parse_pct_val(_at(12)),
        }
    return {}


def _parse_pct_val(v):
    if v in (None, "", "—"):
        return None
    try:
        s = str(v).replace("%", "").replace(",", "").strip()
        return float(s)
    except (ValueError, TypeError):
        return None


def render_overall_hero(portfolio: dict) -> str:
    o = extract_overall(portfolio)
    if not o or o.get("value") is None:
        return ""
    value = o["value"]
    total_change = o.get("total_change")
    inv_change = o.get("inv_change")
    cap_roi = o.get("cap_roi")
    dividends = o.get("dividends")
    change_cls = "pos" if (total_change or 0) >= 0 else "neg"
    change_arrow = "▲" if (total_change or 0) >= 0 else "▼"
    change_sign = "+" if (total_change or 0) >= 0 else ""
    return f"""
<div class="card" style="background:linear-gradient(135deg, var(--bg-card), var(--bg-sunken));border:1px solid var(--ink);margin-bottom:18px;padding:24px 28px;">
  <div style="display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:20px;">
    <div>
      <div class="kpi-label">Overall Portfolio Value</div>
      <div class="kpi-value" style="font-size:2.6rem;line-height:1;margin-top:6px;">{fmt_gbp(value)}</div>
      <div class="kpi-sub" style="margin-top:8px;">
        <span class="{change_cls}">{change_arrow} {change_sign}{fmt_gbp(total_change)}</span>
        {f'&nbsp;<span style="color:var(--muted)">({cap_roi:+.1f}% cap ROI)</span>' if cap_roi is not None else ''}
        &nbsp;<span style="color:var(--muted)">total change since invested</span>
      </div>
    </div>
    <div style="text-align:right;">
      <div class="kpi-label">Invested</div>
      <div class="kpi-value sm" style="font-size:1.3rem;">{fmt_gbp(o.get('invested'))}</div>
      {f'<div class="kpi-sub"><span class="pos">+{fmt_gbp(inv_change)}</span> inv change</div>' if inv_change is not None else ''}
      {f'<div class="kpi-sub" style="margin-top:6px;">Dividends {fmt_gbp(dividends)}</div>' if dividends else ''}
    </div>
  </div>
</div>
"""


def render_liabilities(monthly: dict, school_fees: dict | None = None) -> str:
    """Liabilities breakdown card reading the Monthly tab.

    Pulls latest value + month-over-month + YTD change for the rows we can
    identify: Mortgage, Credit Card, School (tracked separately as committed).
    """
    rows = (monthly or {}).get("rows") or []
    header = (monthly or {}).get("header") or []
    if not rows or not header:
        return ""

    # Map the meaningful trailing columns by their header label.
    def _col(label: str) -> int | None:
        for i, h in enumerate(header):
            if (h or "").strip().lower() == label.lower():
                return i
        return None

    col_current = _col("Current")
    if col_current is None:
        return ""

    # Derive MoM from the previous non-empty month value, YTD from the last
    # "31/12/YYYY" column before Current (the latest year-end snapshot).
    # Scan backwards from col_current to find the previous month's value.
    def _prev_month_col(r: list) -> int | None:
        for i in range(col_current - 1, 0, -1):
            v = r[i] if i < len(r) else None
            if v and str(v).strip():
                return i
        return None

    # Find last "December" column before Current as the YTD baseline.
    # Header columns alternate year-end snapshots and monthly names.
    def _last_december_col(r: list) -> int | None:
        """Scan backwards from col_current-1 for the last December entry."""
        latest = None
        for i in range(col_current - 1, 0, -1):
            h = header[i] if i < len(header) else ""
            if (h or "").strip().lower() == "december":
                v = r[i] if i < len(r) else None
                if v and str(v).strip():
                    return i
        return latest

    def _row_metrics(r: list) -> dict | None:
        """Extract current / mo / yr from a single row using the column structure."""
        cur = parse_money_val(r[col_current] if col_current < len(r) else None)
        if cur is None:
            return None
        prev_i = _prev_month_col(r)
        prev = parse_money_val(r[prev_i]) if prev_i is not None else None
        dec_i = _last_december_col(r)
        ye = parse_money_val(r[dec_i]) if dec_i is not None else None
        mo = (cur - prev) if cur is not None and prev is not None else None
        yr = (cur - ye) if cur is not None and ye is not None else None
        return {"current": cur, "mo": mo, "yr": yr}

    def _extract(label: str, col_idx: int = 0) -> dict | None:
        """Find first row where the cell at col_idx matches label (exact, case-insensitive)."""
        for r in rows:
            if not r or col_idx >= len(r):
                continue
            cell = r[col_idx]
            if not isinstance(cell, str):
                continue
            if cell.strip().lower() != label.lower():
                continue
            return _row_metrics(r)
        return None

    def _extract_sum(substr: str, col_idx: int = 2) -> dict | None:
        """Sum metrics across all rows whose cell at col_idx CONTAINS substr.
        Used for multi-row categories like 'Loans' (currently just MKS Loan,
        but future-proofed for additional loan items)."""
        agg = {"current": 0.0, "mo": 0.0, "yr": 0.0, "count": 0}
        any_data = False
        for r in rows:
            if not r or col_idx >= len(r):
                continue
            cell = r[col_idx]
            if not isinstance(cell, str) or substr.lower() not in cell.lower():
                continue
            m = _row_metrics(r)
            if m is None:
                continue
            agg["current"] += (m["current"] or 0)
            agg["mo"]      += (m["mo"] or 0)
            agg["yr"]      += (m["yr"] or 0)
            agg["count"]   += 1
            any_data = True
        return agg if any_data else None

    mortgage = _extract("Mortgage")
    cc_raw   = _extract("Credit Card")            # sheet aggregate — INCLUDES loans
    loans    = _extract_sum("loan", col_idx=2)    # row 32 = "MKS Loan", future-proofed
    school   = _extract("School")

    # The sheet's "Credit Card" aggregate (row 51) sums per-item rows 28-32
    # which INCLUDES MKS Loan. To show Loans separately without double-counting,
    # subtract the loans figures from the CC aggregate so the displayed CC is
    # pure credit-card debt and Mortgage + Loans + CC === sheet's Total Liabs.
    if cc_raw and loans:
        cc = {
            "current": (cc_raw.get("current") or 0) - (loans.get("current") or 0),
            "mo":      (cc_raw.get("mo")      or 0) - (loans.get("mo")      or 0),
            "yr":      (cc_raw.get("yr")      or 0) - (loans.get("yr")      or 0),
        }
    else:
        cc = cc_raw

    if not mortgage and not cc and not loans:
        return ""

    liab_total = ((mortgage or {}).get("current", 0)
                  + (cc or {}).get("current", 0)
                  + (loans or {}).get("current", 0))
    liab_mo    = (((mortgage or {}).get("mo") or 0)
                  + ((cc or {}).get("mo") or 0)
                  + ((loans or {}).get("mo") or 0))
    liab_yr    = (((mortgage or {}).get("yr") or 0)
                  + ((cc or {}).get("yr") or 0)
                  + ((loans or {}).get("yr") or 0))

    def _row(label: str, d: dict | None, accent: str) -> str:
        if not d:
            return ""
        cur = d.get("current") or 0
        mo = d.get("mo") or 0
        yr = d.get("yr") or 0
        mo_cls = "pos" if mo <= 0 else "neg"  # paying down = good
        yr_cls = "pos" if yr <= 0 else "neg"
        mo_sign = "+" if mo > 0 else ""
        yr_sign = "+" if yr > 0 else ""
        pct = (cur / liab_total * 100) if liab_total else 0
        return f"""
      <tr>
        <td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{accent};margin-right:7px;"></span>{label}</td>
        <td class="num">{fmt_gbp(cur)}</td>
        <td class="num">{pct:.1f}%</td>
        <td class="num {mo_cls}">{mo_sign}{fmt_gbp(mo)}</td>
        <td class="num {yr_cls}">{yr_sign}{fmt_gbp(yr)}</td>
      </tr>"""

    liab_rows = (
        _row("Mortgage",    mortgage, "#ef4444")
        + _row("Loans",       loans,  "#a855f7")
        + _row("Credit Card", cc,     "#f59e0b")
    )

    total_mo_cls = "pos" if liab_mo <= 0 else "neg"
    total_yr_cls = "pos" if liab_yr <= 0 else "neg"
    total_mo_sign = "+" if liab_mo > 0 else ""
    total_yr_sign = "+" if liab_yr > 0 else ""

    # Stacked bar visualising mortgage vs loans vs CC proportions
    mort_pct  = ((mortgage or {}).get("current", 0) / liab_total * 100) if liab_total else 0
    loans_pct = ((loans    or {}).get("current", 0) / liab_total * 100) if liab_total else 0
    cc_pct    = ((cc       or {}).get("current", 0) / liab_total * 100) if liab_total else 0

    committed_section = ""
    # Show the row if we have EITHER the Monthly-tab School row OR the School Fees
    # summary dict — the latter is authoritative for YTD-paid.
    sf = school_fees or {}
    if school or sf:
        # Remaining to pay: prefer the explicit (total_cost - paid) calc; fall back to
        # the Monthly tab's "School" current-balance cell.
        total_cost = sf.get("total_cost") or 0
        paid_to_date = sf.get("paid") or 0
        remaining_to_pay = (total_cost - paid_to_date) if total_cost else ((school or {}).get("current") or 0)

        # MoM: latest month's payment from the Monthly-tab "School" row delta
        # (already a negative number when the remaining balance dropped).
        s_mo = (school or {}).get("mo") or 0

        # YTD: the authoritative figure is the sum of column S for the current year
        # (school_fees.paid_ytd) — this is actual cash paid YTD, regardless of
        # tuition-increase accruals. Show as negative (money out).
        paid_ytd = sf.get("paid_ytd")
        if paid_ytd is not None:
            s_yr = -paid_ytd
        else:
            s_yr = (school or {}).get("yr") or 0

        s_mo_cls = "pos" if s_mo <= 0 else "neg"
        s_yr_cls = "pos" if s_yr <= 0 else "neg"
        s_mo_sign = "+" if s_mo > 0 else ""
        s_yr_sign = "+" if s_yr > 0 else ""
        ytd_year = sf.get("ytd_year") or ""
        ytd_header = f"YTD {ytd_year}" if ytd_year else "YTD"
        committed_section = f"""
  <h3 style="font-size:13px;margin-top:20px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;">Committed Future Spend</h3>
  <table>
    <thead>
      <tr>
        <th>Item</th>
        <th class="num">Remaining</th>
        <th class="num">MoM</th>
        <th class="num">{ytd_header}</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#8b5cf6;margin-right:7px;"></span>School Fees</td>
        <td class="num">{fmt_gbp(remaining_to_pay)}</td>
        <td class="num {s_mo_cls}">{s_mo_sign}{fmt_gbp(s_mo)}</td>
        <td class="num {s_yr_cls}">{s_yr_sign}{fmt_gbp(s_yr)}</td>
      </tr>
    </tbody>
  </table>
  <small style="color:var(--muted);">Tracked outside liabilities — YTD = actual payments this year (col S of schedule).</small>
"""

    return f"""
<div class="card">
  <h2>Liabilities Breakdown</h2>
  <div style="display:flex;gap:12px;margin-bottom:12px;align-items:baseline;">
    <div>
      <div class="kpi-label">Total Liabilities</div>
      <div class="kpi-value" style="color:var(--accent);">{fmt_gbp(liab_total)}</div>
    </div>
    <div style="font-family:var(--f-mono);font-size:12px;color:var(--muted);">
      <span class="{total_mo_cls}">{total_mo_sign}{fmt_gbp(liab_mo)}</span> MoM ·
      <span class="{total_yr_cls}">{total_yr_sign}{fmt_gbp(liab_yr)}</span> YTD
    </div>
  </div>
  <div class="stacked-bar" style="margin-bottom:12px;">
    <div class="seg neg"  style="width:{mort_pct:.1f}%;background:#ef4444;">Mortgage {mort_pct:.0f}%</div>
    <div class="seg"      style="width:{loans_pct:.1f}%;background:#a855f7;color:#fff;">Loans {loans_pct:.0f}%</div>
    <div class="seg warn" style="width:{cc_pct:.1f}%;background:#f59e0b;">CC {cc_pct:.0f}%</div>
  </div>
  <table>
    <thead>
      <tr>
        <th>Category</th>
        <th class="num">Balance</th>
        <th class="num">% of total</th>
        <th class="num">MoM</th>
        <th class="num">YTD</th>
      </tr>
    </thead>
    <tbody>{liab_rows}
      <tr style="border-top:2px solid var(--border);font-weight:600;">
        <td>Total</td>
        <td class="num">{fmt_gbp(liab_total)}</td>
        <td class="num">100%</td>
        <td class="num {total_mo_cls}">{total_mo_sign}{fmt_gbp(liab_mo)}</td>
        <td class="num {total_yr_cls}">{total_yr_sign}{fmt_gbp(liab_yr)}</td>
      </tr>
    </tbody>
  </table>
  <small style="color:var(--muted);">Negative MoM/YTD = paying down (good). Source: Monthly tab.</small>
  {committed_section}
</div>
"""


def render_portfolio_raw(portfolio: dict) -> str:
    """Render the Portfolio tab exactly as it appears in the Google Sheet.
    Detects section headers (col B == 'Investment'), totals (col B starts with 'Total'),
    and empty spacer rows.
    """
    rows = (portfolio or {}).get("rows") or []
    if not rows:
        return ""

    # Trim everything BELOW 'Overall' (the block after it is duplicated noise),
    # but KEEP the Overall row itself as a grand-total footer.
    overall_idx = None
    for i, r in enumerate(rows):
        col_a = (r[0] if len(r) > 0 else "").strip().lower() if r else ""
        if col_a == "overall":
            overall_idx = i
            break
    if overall_idx is not None:
        rows = rows[:overall_idx + 1]
    # Drop trailing blank rows after truncation
    while rows and not any((c or "").strip() for c in rows[-1]):
        rows.pop()

    # Which columns to show (0-indexed). Portfolio layout:
    # 0 Account | 1 Investment | 2 Type | 3 (prior value, noisy) | 4 Ticker | 5 Units Held
    # 6 Current Value | 7 Invested | 8 Inv Change | 9 Total Change | 10 Cap ROI
    # 11 Dividends | 12 Div Yield
    COLS = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    HEADERS = ["Account", "Investment", "Type", "Ticker", "Units", "Current Value",
               "Invested", "Inv Change", "Total Change", "Cap ROI", "Dividends", "Div Yield"]
    NUMERIC_COLS = {5, 6, 7, 8, 9, 10, 11, 12}  # format these right-aligned
    # Source columns where the sign drives colour (reds/greens)
    SIGNED_COLS = {8, 9, 10}  # Inv Change, Total Change, Cap ROI
    CURRENT_VALUE_COL = 6

    def _is_negative(val) -> bool | None:
        """Return True if value is numerically negative, False if positive,
        None if not numeric / empty."""
        if val is None:
            return None
        s = str(val).strip()
        if not s:
            return None
        # Strip currency, percentage, thousands separators
        cleaned = s.replace("£", "").replace("$", "").replace(",", "").replace("%", "").strip()
        if not cleaned or cleaned in ("-", "—"):
            return None
        try:
            return float(cleaned) < 0
        except (ValueError, TypeError):
            return None

    # AJ Bell Neil sub-group mapping (from the AI tab's AI / Defense / UK buckets).
    # Each row inside "AJ Bell Neil" is classified by investment name into one of 3 groups.
    # Matches AI tab summary totals (AI £41,190, Defense £10,867) exactly.
    AJB_AI_NAMES = {
        "palantir", "robinhood", "roblox", "roku", "shopify", "tempus ai",
        "tesla", "unity software", "crispr therapeutics", "teradyne",
        "bitmine immersion", "advanced micro devices", "alibaba",
    }
    AJB_DEFENSE_NAMES = {
        "kratos defense", "rocket lab", "archer aviation",
        "iridium communications", "trimble",
    }

    def _ajb_group(name: str) -> str:
        n = name.lower().strip()
        for key in AJB_AI_NAMES:
            if key in n:
                return "ai"
        for key in AJB_DEFENSE_NAMES:
            if key in n:
                return "defense"
        return "ukother"

    AJB_LABELS = {
        "ai":      ("🤖 AI / Growth Tech",     "#4f46e5"),
        "defense": ("🛡️ Defense & Aerospace",  "#b8472c"),
        "ukother": ("🇬🇧 UK / Other",           "#2d5b3e"),
    }

    # First pass: detect AJ Bell Neil section boundaries and group its rows.
    # We process rows sequentially, but inside AJ Bell Neil we bucket first
    # then emit in order AI → Defense → UK/Other with sub-headers + dividers.
    def _render_single_row(row, col_b, col_a, is_section_header, is_total, is_grand_total):
        colour_signed = not is_section_header
        cells_html = []
        for ci in COLS:
            val = row[ci] if ci < len(row) else ""
            val_str = str(val) if val is not None else ""
            cls_parts = []
            if ci in NUMERIC_COLS:
                cls_parts.append("num")
            if ci == CURRENT_VALUE_COL and not is_section_header:
                cls_parts.append("portfolio-current")
            if ci in SIGNED_COLS and colour_signed:
                neg = _is_negative(val)
                if neg is True:
                    cls_parts.append("neg")
                elif neg is False:
                    cls_parts.append("pos")
            cls = f' class="{" ".join(cls_parts)}"' if cls_parts else ""
            cells_html.append(f"<td{cls}>{val_str}</td>")
        if is_grand_total:
            tr_class = "portfolio-grand-total"
        elif is_section_header:
            tr_class = "portfolio-header"
        elif is_total:
            tr_class = "portfolio-total"
        else:
            tr_class = ""
        tr_open = f'<tr class="{tr_class}">' if tr_class else "<tr>"
        return tr_open + "".join(cells_html) + "</tr>"

    html_rows = []
    in_ajb = False          # inside "AJ Bell Neil" section (not ISA)
    ajb_buckets: dict[str, list] = {"ai": [], "defense": [], "ukother": []}

    def _flush_ajb():
        """Emit grouped AJ Bell Neil rows with thin dividers + sub-headers."""
        if not any(ajb_buckets.values()):
            return
        first_group = True
        for gk in ("ai", "defense", "ukother"):
            group_rows = ajb_buckets[gk]
            if not group_rows:
                continue
            label, colour = AJB_LABELS[gk]
            # Thin divider + sub-header row
            divider_cls = "portfolio-subgroup" + ("" if first_group else " portfolio-subgroup-divider")
            html_rows.append(
                f'<tr class="{divider_cls}"><td colspan="{len(COLS)}" '
                f'style="border-top:1px solid {colour}33;padding:8px 6px;'
                f'font-size:11px;font-weight:600;color:{colour};'
                f'text-transform:uppercase;letter-spacing:0.06em;'
                f'background:linear-gradient(to right, {colour}0D, transparent);">'
                f'{label}</td></tr>'
            )
            for entry in group_rows:
                html_rows.append(_render_single_row(*entry))
            first_group = False
        ajb_buckets["ai"].clear()
        ajb_buckets["defense"].clear()
        ajb_buckets["ukother"].clear()

    for row in rows:
        col_b = (row[1] if len(row) > 1 else "").strip() if row else ""
        is_blank = not any((c or "").strip() for c in row)

        if is_blank:
            # If we were in AJB section, flush first.
            if in_ajb:
                _flush_ajb()
                in_ajb = False
            html_rows.append(f'<tr class="portfolio-spacer"><td colspan="{len(COLS)}">&nbsp;</td></tr>')
            continue

        col_a = (row[0] if len(row) > 0 else "").strip() if row else ""
        is_section_header = col_b.lower() == "investment"
        is_total = col_b.lower().startswith("total ")
        is_grand_total = col_a.lower() == "overall"

        # Detect AJ Bell Neil section boundary (but NOT the ISA).
        if is_section_header:
            if in_ajb:
                _flush_ajb()
                in_ajb = False
            acc_name = col_a.lower()
            if acc_name == "aj bell neil":  # exact match — excludes "AJ Bell Neil ISA"
                in_ajb = True
        elif is_total and in_ajb and "aj bell neil" in col_b.lower():
            _flush_ajb()
            in_ajb = False
        elif in_ajb and not is_section_header and not is_total and not is_grand_total:
            # Bucket this row into AI / Defense / UK-Other
            group = _ajb_group(col_b)
            ajb_buckets[group].append((row, col_b, col_a, is_section_header, is_total, is_grand_total))
            continue  # deferred — emit via _flush_ajb()

        # Normal emit for rows outside AJB or for headers/totals inside AJB.
        html_rows.append(_render_single_row(row, col_b, col_a, is_section_header, is_total, is_grand_total))

    # End-of-rows safety flush.
    if in_ajb:
        _flush_ajb()

    # thead uses our own headers (not the sheet's repeating header rows)
    thead = (
        "<tr>"
        + "".join(
            f'<th class="num">{h}</th>' if i in NUMERIC_COLS else f'<th>{h}</th>'
            for i, h in zip(COLS, HEADERS)
        )
        + "</tr>"
    )

    return f"""
<div class="card" style="overflow-x:auto">
  <h2>Portfolio</h2>
  <small style="color:var(--muted)">Mirror of the Portfolio tab in Finances ND.</small>
  <table class="portfolio-table">
    <thead>{thead}</thead>
    <tbody>{"".join(html_rows)}</tbody>
  </table>
</div>
"""


def render_investments_chart(inv: dict) -> str:
    """Pension vs Non-Pension over years — stacked bar with a total line."""
    years = inv.get("years") or []
    series = inv.get("series") or {}
    pension = series.get("Pension") or []
    non_pension = series.get("Non-Pension") or []
    # Compute total & non-pension %
    totals = []
    non_pens_pct = []
    for i, y in enumerate(years):
        p = pension[i] if i < len(pension) else None
        np_ = non_pension[i] if i < len(non_pension) else None
        if p is not None and np_ is not None:
            tot = p + np_
            totals.append(tot)
            non_pens_pct.append(round(np_ / tot * 100, 1) if tot else 0)
        else:
            totals.append(None)
            non_pens_pct.append(None)

    rows = []
    for i, y in enumerate(years):
        p = pension[i] if i < len(pension) else None
        np_ = non_pension[i] if i < len(non_pension) else None
        t = totals[i]
        rows.append(
            f'<tr><td>{y}</td>'
            f'<td class="num">{fmt_gbp(p)}</td>'
            f'<td class="num">{fmt_gbp(np_)}</td>'
            f'<td class="num"><strong>{fmt_gbp(t)}</strong></td>'
            f'<td class="num" style="color:var(--text-2)">{(non_pens_pct[i] if non_pens_pct[i] is not None else 0):.0f}%</td>'
            f'</tr>'
        )

    # Add a YoY growth column: change in Total from prior year
    rows2 = []
    for i, y in enumerate(years):
        p = pension[i] if i < len(pension) else None
        np_ = non_pension[i] if i < len(non_pension) else None
        t = totals[i]
        prev_t = totals[i-1] if i > 0 else None
        growth = (t - prev_t) if (t is not None and prev_t is not None) else None
        growth_pct = (growth / prev_t * 100) if growth is not None and prev_t else None
        rows2.append(
            f'<tr>'
            f'<td>{y}</td>'
            f'<td class="num">{fmt_gbp(p)}</td>'
            f'<td class="num">{fmt_gbp(np_)}</td>'
            f'<td class="num"><strong>{fmt_gbp(t)}</strong></td>'
            f'<td class="num" style="color:var(--text-2)">{(non_pens_pct[i] if non_pens_pct[i] is not None else 0):.0f}%</td>'
            f'<td class="num {colour_cls(growth)}">{fmt_gbp(growth) if growth is not None else "—"}</td>'
            f'<td class="num {colour_cls(growth_pct)}">{fmt_pct(growth_pct)}</td>'
            f'</tr>'
        )

    return f"""
<div class="card">
  <h2>Pension vs Non-Pension</h2>
  <table>
    <thead>
      <tr>
        <th>Year</th>
        <th class="num">Pension</th>
        <th class="num">Non-Pension</th>
        <th class="num">Total</th>
        <th class="num">Non-Pen %</th>
        <th class="num">YoY Change</th>
        <th class="num">YoY %</th>
      </tr>
    </thead>
    <tbody>{''.join(rows2)}</tbody>
  </table>
</div>
"""


# ── Main ──────────────────────────────────────────────────────────────────────
def build_html(data: dict) -> str:
    nw = data["net_worth"]
    cur_net = (nw.get("current") or {}).get("net")

    sections = [
        render_headline(nw, data.get("investments") or {}),
        render_yoy_chart(nw),
        render_investments_chart(data.get("investments") or {}),
        render_allocation(
            data.get("portfolio") or {},
            (data.get("retirement") or {}).get("wealth") or {},
            nw.get("current") or {},
        ),
        render_liabilities(data.get("monthly") or {}, data.get("school_fees") or {}),
        render_retirement(data.get("retirement") or {}, cur_net),
        render_school_fees(data.get("school_fees") or {}),
        render_portfolio_raw(data.get("portfolio") or {}),
    ]

    extra_css = """
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin: 16px 0; }
    .kpi { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; padding: 14px; }
    .kpi-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted); margin-bottom: 4px; }
    .kpi-value { font-size: 22px; font-weight: 600; font-variant-numeric: tabular-nums; }
    .kpi-value.sm { font-size: 17px; }
    .kpi-sub { font-size: 12px; margin-top: 2px; }
    .progress { height: 10px; background: var(--bg-2, #1e2330); border-radius: 5px; overflow: hidden; }
    .bar { height: 100%; background: var(--accent); transition: width 0.5s; }
    .bar.pos { background: #10b981; }
    .bar.warn { background: #f59e0b; }
    .bar.neg { background: #ef4444; }
    .stacked-bar { display:flex; height:28px; border-radius:6px; overflow:hidden; font-size:11px; font-weight:600; color:white; }
    .stacked-bar .seg { display:flex; align-items:center; justify-content:center; padding: 0 6px; }
    .stacked-bar .seg.pos { background:#10b981; }
    .stacked-bar .seg.warn { background:#f59e0b; }
    .stacked-bar .seg.neg { background:#ef4444; }
    .pos { color: #10b981; }
    .neg { color: #ef4444; }
    .warn { color: #f59e0b; }
    table.tight td, table.tight th { padding: 4px 8px; font-size: 13px; }
    /* Portfolio-tab mirror */
    .portfolio-table { font-size: 13px; }
    .portfolio-table td, .portfolio-table th { padding: 5px 8px; }
    .portfolio-table .portfolio-header td {
      background: rgba(79, 70, 229, 0.12);
      color: var(--accent, #4f46e5);
      font-weight: 700;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.03em;
      border-top: 2px solid var(--accent, #4f46e5);
      border-bottom: 1px solid var(--border);
    }
    .portfolio-table .portfolio-total td {
      font-weight: 600;
      border-top: 1px solid var(--border);
      background: rgba(16, 185, 129, 0.07);
    }
    .portfolio-table .portfolio-spacer td { padding: 3px 0; border: none; background: transparent; }
    .portfolio-table .portfolio-grand-total td {
      font-weight: 700;
      font-size: 14px;
      background: rgba(79, 70, 229, 0.18);
      color: var(--accent, #4f46e5);
      border-top: 2px solid var(--accent, #4f46e5);
      border-bottom: 2px solid var(--accent, #4f46e5);
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }
    /* Highlight the Current Value column so it jumps out */
    .portfolio-table td.portfolio-current {
      font-weight: 700;
      color: var(--ink, #1a1a1a);
      background: rgba(16, 185, 129, 0.05);
    }
    .portfolio-table .portfolio-header td.portfolio-current,
    .portfolio-table .portfolio-total  td.portfolio-current,
    .portfolio-table .portfolio-grand-total td.portfolio-current {
      background: rgba(16, 185, 129, 0.14);
    }
    /* Green/red for signed change columns */
    .portfolio-table td.pos { color: #10b981; font-weight: 500; }
    .portfolio-table td.neg { color: #ef4444; font-weight: 500; }
    /* Keep grand-total colour inheritance: overall row should still read as accent */
    .portfolio-table .portfolio-grand-total td.pos,
    .portfolio-table .portfolio-grand-total td.neg {
      font-weight: 700;
    }
    """

    generated = data.get("generated_at", datetime.now().strftime("%Y-%m-%d %H:%M"))
    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Finances Dashboard</title>
{_FONTS_LINK}
<style>{_EDITORIAL_CSS}{extra_css}</style>
</head><body>
{_nav_html("finances", privacy=True)}
<main class="container">
  <h1>Finances Dashboard</h1>
  <small style="color:var(--muted)">Source: Google Sheets "Finances ND" • Generated {generated}</small>
  {"".join(sections)}
</main>
{_THEME_JS}
</body></html>
"""


def main() -> None:
    if not DATA.exists():
        import sys
        sys.exit(f"Missing {DATA}. Run: python scripts/sync_finances.py first.")
    data = json.loads(DATA.read_text(encoding="utf-8"))
    html = build_html(data)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"Dashboard saved -> {OUTPUT}")


if __name__ == "__main__":
    main()
