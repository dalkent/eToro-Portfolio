"""sync_finances.py — Pull data from the 'Finances ND' Google Sheet into
data/finances.json, structured for the Finances dashboard.

Source tabs used:  Dashboard, Portfolio, Crypto, Retirement Planning, Monthly.
Tabs skipped:      Divorce Calc (privacy).
"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from google_auth import get_credentials  # noqa: E402

try:
    from googleapiclient.discovery import build
except ImportError:
    sys.exit("Install: pip install google-api-python-client")

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT = DATA_DIR / "finances.json"

SHEET_ID = "1fo3DMv0YqeOlMA4zMIgrWkS-eeyI9CwJ3WRZMhtryK8"

# ── Parsing helpers ────────────────────────────────────────────────────────────
_MONEY_RE = re.compile(r"[£$,\s]")


def parse_money(v) -> float | None:
    """Parse '£1,234.56', '-£12', '$1,200' → float. Empty/None → None."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in {"-", "—", "N/A"}:
        return None
    # Handle minus signs
    neg = s.startswith("-") or s.startswith("(-")
    s = _MONEY_RE.sub("", s.lstrip("-(").rstrip(")"))
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


def parse_pct(v) -> float | None:
    """Parse '12.5%' → 12.5."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v) * 100  # Google stores 0.125 for 12.5%
    s = str(v).strip().rstrip("%").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


# ── Sheet client ───────────────────────────────────────────────────────────────
def get_values(sheets_api, range_: str) -> list[list]:
    resp = sheets_api.values().get(
        spreadsheetId=SHEET_ID, range=range_, valueRenderOption="FORMATTED_VALUE"
    ).execute()
    return resp.get("values", [])


# ── Section builders ──────────────────────────────────────────────────────────
def build_net_worth(sheets_api) -> dict:
    """Dashboard!A1:H5 — Total Assets, Liabilities, Net Assets by year + current."""
    rows = get_values(sheets_api, "Dashboard!A1:H5")
    # Row 0: ["Year", "2020", "2021", "2022", "2023", "2024", "2025", "Current"]
    years = rows[0][1:]
    def row_vals(r):
        return [parse_money(x) for x in r[1:]]

    assets      = row_vals(rows[1]) if len(rows) > 1 else []
    liabilities = row_vals(rows[2]) if len(rows) > 2 else []
    net         = row_vals(rows[3]) if len(rows) > 3 else []

    history = []
    for i, y in enumerate(years):
        a = assets[i] if i < len(assets) else None
        l = liabilities[i] if i < len(liabilities) else None
        n = net[i] if i < len(net) else None
        history.append({"year": y, "assets": a, "liabilities": l, "net": n})

    # yoy change from the 'Change' row (row index 4, col 2 onward since 2020 has no yoy)
    changes = row_vals(rows[4]) if len(rows) > 4 else []
    for i, h in enumerate(history):
        h["change"] = changes[i] if i < len(changes) else None

    # Current
    current = history[-1] if history else None

    return {"history": history, "current": current}


def build_investments_breakdown(sheets_api) -> dict:
    """Dashboard!A8:H14 — Pension vs Non-Pension over years."""
    rows = get_values(sheets_api, "Dashboard!A8:H14")
    # Row 0 (index 0): ["Year","2020",...,"Current"]
    years = rows[0][1:] if rows else []
    breakdown = {"years": years, "series": {}}
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        label = row[0].strip()
        if label and label.lower() != "year":
            breakdown["series"][label] = [parse_money(v) for v in row[1:]]
    return breakdown


def build_school_fees(sheets_api) -> dict:
    """Dashboard!J1:M10 — School Fees summary.
    Dashboard!O1:S200 — monthly payment schedule (Date / Fee / Extra / Note / Total Paid).
    """
    rows = get_values(sheets_api, "Dashboard!J1:M10")
    data = {}
    for row in rows:
        if not row:
            continue
        key = (row[0] or "").strip()
        if not key or key.lower() == "school fees":
            continue
        val = row[1] if len(row) > 1 else None
        if "years remaining" in key.lower():
            data["years_remaining"] = parse_money(val)
        elif "approx cost" in key.lower():
            data["total_cost"] = parse_money(val)
        elif key.lower() == "paid":
            data["paid"] = parse_money(val)
        elif key.lower() == "savings":
            data["saved"] = parse_money(val)
        elif key.lower() == "remaining":
            data["remaining"] = parse_money(val)
        elif "predicted savings" in key.lower():
            data["predicted_savings"] = parse_money(val)
        elif "new remaining" in key.lower():
            data["new_remaining"] = parse_money(val)
    # percentages
    total = data.get("total_cost") or 0
    if total:
        data["paid_pct"] = (data.get("paid") or 0) / total * 100
        data["saved_pct"] = (data.get("saved") or 0) / total * 100
        data["remaining_pct"] = (data.get("remaining") or 0) / total * 100

    # Monthly schedule → Paid YTD (sum of col S for rows in the current calendar year).
    schedule_rows = get_values(sheets_api, "Dashboard!O1:S200")
    paid_ytd = 0.0
    ytd_months = 0
    today = datetime.now().date()
    year_str = str(today.year)
    short_year = year_str[-2:]
    for sr in schedule_rows[1:]:  # skip header
        if not sr or len(sr) < 5:
            continue
        date_cell = (sr[0] or "").strip() if sr[0] else ""
        total_cell = sr[4] if len(sr) > 4 else None
        amt = parse_money(total_cell)
        if amt is None or amt == 0:
            continue
        # Date comes as "01/01/26" / "1/3/26" — match current year from the trailing token.
        tokens = re.split(r"[/\-\.]", date_cell)
        if len(tokens) < 3:
            continue
        y = tokens[-1].strip()
        if y == year_str or y == short_year:
            paid_ytd += amt
            ytd_months += 1
    data["paid_ytd"] = paid_ytd
    data["ytd_months"] = ytd_months
    data["ytd_year"] = today.year
    return data


def build_retirement(sheets_api) -> dict:
    """Retirement Planning tab — readiness table (cols B/C/D, rows 9-30) and
    pot projections by retirement age (Q4:T33)."""

    # ── Readiness scenarios: cols B, C, D with key rows ───────────────────────
    rows = get_values(sheets_api, "Retirement Planning!A1:D33")
    def cell(r: int, c: int):
        """1-based row/col, returns raw string or None."""
        if r - 1 >= len(rows):
            return None
        row = rows[r - 1]
        if c - 1 >= len(row):
            return None
        return row[c - 1]

    scenarios = []
    for col_idx in (2, 3, 4):  # B, C, D
        monthly = parse_money(cell(9, col_idx))
        yearly = parse_money(cell(10, col_idx))
        rate = parse_pct(cell(12, col_idx))
        pot_required = parse_money(cell(13, col_idx))
        current_inv = parse_money(cell(15, col_idx))
        gap = parse_money(cell(16, col_idx))
        house_total = parse_money(cell(19, col_idx))
        gap_with_house = parse_money(cell(20, col_idx))
        years_no_house = parse_money(cell(28, col_idx))
        years_with_house = parse_money(cell(29, col_idx))
        months = parse_money(cell(30, col_idx))
        if monthly is None or monthly == 0:
            continue
        scenarios.append({
            "col": chr(ord('A') + col_idx - 1),
            "monthly": monthly,
            "yearly": yearly,
            "spend_rate_pct": rate,
            "pot_required": pot_required,
            "current_investments": current_inv,
            "gap": gap,
            "total_with_house": house_total,
            "gap_with_house": gap_with_house,
            "years_no_house": years_no_house,
            "years_with_house": years_with_house,
            "months_to_save": months,
        })

    # ── Pot projections at ages 55 / 57 / 60 (columns Q..T, rows 4-33) ───────
    pblock = get_values(sheets_api, "Retirement Planning!Q1:T33")
    def pcell(r: int, c: int):
        if r - 1 >= len(pblock):
            return None
        row = pblock[r - 1]
        if c - 1 >= len(row):
            return None
        return row[c - 1]

    # Block starts: age 60 → row 4-11, age 57 → 15-22, age 55 → 26-33
    def extract_block(start_row: int) -> dict:
        # Layout relative to start_row:
        # +0: "Current Pot" | value
        # +1: "Retirement Date" | date
        # +2: "Age at retirement" | age
        # +3: "Months remaining" | months
        # +5: "Expected return" | 5% | 8% | 10%
        # +6: "Pot if continue invest" | val | val | val
        # +7: "Pot no more invest from today" | val | val | val
        return {
            "current_pot": parse_money(pcell(start_row + 0, 2)),
            "retirement_date": pcell(start_row + 1, 2),
            "age": parse_money(pcell(start_row + 2, 2)),
            "months_remaining": parse_money(pcell(start_row + 3, 2)),
            "return_rates": [pcell(start_row + 5, c) for c in (2, 3, 4)],
            "pot_continue_invest": [parse_money(pcell(start_row + 6, c)) for c in (2, 3, 4)],
            "pot_stop_invest":    [parse_money(pcell(start_row + 7, c)) for c in (2, 3, 4)],
        }

    projections = [
        {"label": "Retire at 60", **extract_block(4)},
        {"label": "Retire at 57", **extract_block(15)},
        {"label": "Retire at 55", **extract_block(26)},
    ]
    # The age-55 block in the sheet omits the Current Pot value (implied same as others).
    # Fill any missing current_pot from the first block that has one.
    seed_pot = next((p["current_pot"] for p in projections if p.get("current_pot")), None)
    for p in projections:
        if p.get("current_pot") is None:
            p["current_pot"] = seed_pot

    return {"scenarios": scenarios, "projections": projections}


def build_portfolio_raw(sheets_api) -> dict:
    """Portfolio tab — pull the raw grid so we can render it 'as is'."""
    rows = get_values(sheets_api, "Portfolio!A1:N200")
    # Trim trailing blank rows
    while rows and not any((c or "").strip() for c in rows[-1]):
        rows.pop()
    # Pad every row to consistent width
    width = max((len(r) for r in rows), default=0)
    padded = [(r + [""] * (width - len(r))) for r in rows]
    return {"width": width, "rows": padded}


def build_monthly_series(sheets_api) -> dict:
    """Monthly tab — extract headers + full row grid so the dashboard can
    pull out liabilities categories (CC / Loans / Mortgage) and other series.
    """
    rows = get_values(sheets_api, "Monthly!A1:CI300")
    if not rows:
        return {"months_header": [], "rows": []}
    header = rows[0] if rows else []
    months = [(h or "").strip() for h in header[3:] if h]
    # Normalize row widths
    width = max((len(r) for r in rows), default=0)
    padded = [(r + [""] * (width - len(r))) for r in rows]
    return {
        "months_header": months[:60],
        "header":        [(h or "").strip() for h in header],
        "rows":          padded,
    }


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    sheets_api = service.spreadsheets()

    print("Fetching Finances ND …")
    out = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "net_worth": build_net_worth(sheets_api),
        "investments": build_investments_breakdown(sheets_api),
        "school_fees": build_school_fees(sheets_api),
        "retirement": build_retirement(sheets_api),
        "portfolio": build_portfolio_raw(sheets_api),
        "monthly": build_monthly_series(sheets_api),
    }

    OUTPUT.parent.mkdir(exist_ok=True)
    OUTPUT.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    nw = out["net_worth"]["current"] or {}
    print(f"Saved {OUTPUT}")
    print(f"Current net assets: £{(nw.get('net') or 0):,.0f}")
    print(f"Portfolio rows captured: {len(out['portfolio'].get('rows', []))}")


if __name__ == "__main__":
    main()
