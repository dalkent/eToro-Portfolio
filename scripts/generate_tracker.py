#!/usr/bin/env python3
"""
generate_tracker.py
───────────────────
Generates the weekly FTSE Valuation Tracker article for Substack (paid tier).

Reads eToro_Master.xlsx for:
  - Assumptions sheet: blended targets, signals, model outputs
  - Portfolio sheet: current holdings
  - Tickers sheet: full FTSE universe

Fetches live prices from Yahoo Finance via yfinance.

Outputs:
  - Drafts/YYYY-MM-DD FTSE Valuation Tracker.md  (Substack-ready markdown)

Usage:
  python scripts/generate_tracker.py
  python scripts/generate_tracker.py --date 2026-04-08
  python scripts/generate_tracker.py --no-fetch   (skip Yahoo Finance, use placeholder prices)
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime, date
from collections import defaultdict

import openpyxl

BASE_DIR    = Path(__file__).parent.parent
DATA_DIR    = BASE_DIR / "data"
MASTER      = DATA_DIR / "eToro_Master.xlsx"
DRAFTS_DIR  = Path(__file__).parent.parent.parent / "eToro & Investing" / "Drafts"

# ── Signal logic ─────────────────────────────────────────────────────────────

SIGNAL_ORDER = {
    "Strong Buy": 0, "Buy": 1, "Fair Value": 2, "Sell": 3, "Strong Sell": 4,
    "No Signal": 5, "N/A": 6, "": 6, None: 6,
}

SIGNAL_EMOJI = {
    "Strong Buy": "🟢", "Buy": "🟩", "Fair Value": "🟡",
    "Sell": "🔶", "Strong Sell": "🔴",
}

def compute_signal(vr):
    if vr is None:
        return "N/A"
    if vr >= 1.25:
        return "Strong Buy"
    if vr >= 1.10:
        return "Buy"
    if vr >= 0.90:
        return "Fair Value"
    if vr >= 0.75:
        return "Sell"
    return "Strong Sell"


# ── Load Excel ───────────────────────────────────────────────────────────────

def load_master():
    print(f"Reading {MASTER} ...")
    wb = openpyxl.load_workbook(str(MASTER), data_only=True)

    # GBP/USD rate
    ws_a = wb["Assumptions"]
    gbpusd = 1.34
    for row in ws_a.iter_rows(min_row=3, max_row=6, values_only=True):
        if row[0] == "GBP/USD" and row[1]:
            try:
                gbpusd = float(row[1])
            except Exception:
                pass
            break

    # Portfolio tickers
    ws_p = wb["Portfolio"]
    portfolio_tickers = set()
    for row in ws_p.iter_rows(min_row=3, max_row=300, values_only=True):
        yahoo = str(row[4] or "").strip()
        if yahoo and yahoo.endswith(".L"):
            portfolio_tickers.add(yahoo)

    # Tickers sheet - get Yahoo ticker for each eToro ticker
    ws_t = wb["Tickers"]
    ticker_yahoo = {}
    for row in ws_t.iter_rows(min_row=2, max_row=300, values_only=True):
        etoro = str(row[3] or "").strip()
        yahoo = str(row[5] or "").strip()
        if etoro and yahoo:
            ticker_yahoo[etoro] = yahoo

    # Assumptions - all FTSE valuations
    stocks = []
    for row in ws_a.iter_rows(min_row=7, max_row=300, values_only=True):
        ticker = row[0]
        if not ticker or not isinstance(ticker, str) or not ticker.endswith(".L"):
            continue

        company     = str(row[1] or "").strip()
        sector      = str(row[2] or "").strip()
        beta        = float(row[3]) if row[3] is not None else None
        wacc        = float(row[4]) if row[4] is not None else None
        val1        = float(row[8]) if row[8] is not None else None
        val2        = float(row[9]) if row[9] is not None else None
        val3        = float(row[10]) if row[10] is not None else None
        blended_gbp = float(row[11]) if row[11] is not None else None
        model       = str(row[12] or "").strip()
        updated     = str(row[13] or "").strip()
        prev_signal = str(row[15] or "").strip()
        curr_signal = str(row[16] or "").strip()

        # Convert blended target from GBP to pence
        blended_p = round(blended_gbp * 100, 1) if blended_gbp else None

        in_portfolio = ticker in portfolio_tickers

        stocks.append({
            "ticker": ticker,
            "company": company,
            "sector": sector,
            "beta": beta,
            "wacc": wacc,
            "val1": val1,
            "val2": val2,
            "val3": val3,
            "blended_gbp": blended_gbp,
            "blended_p": blended_p,
            "model": model,
            "updated": updated,
            "prev_signal": prev_signal,
            "curr_signal": curr_signal,
            "in_portfolio": in_portfolio,
            "live_price_p": None,  # populated by fetch
            "value_ratio": None,
            "computed_signal": None,
        })

    return stocks, gbpusd, portfolio_tickers


# ── Fetch live prices ────────────────────────────────────────────────────────

def fetch_prices(stocks):
    """Fetch current prices from Yahoo Finance. Populates live_price_p (in pence)."""
    try:
        import yfinance as yf
    except ImportError:
        print("Warning: yfinance not installed. Run: pip install yfinance")
        return

    tickers = [s["ticker"] for s in stocks if s["blended_p"] is not None]
    print(f"Fetching {len(tickers)} prices from Yahoo Finance ...")

    for s in stocks:
        if s["blended_p"] is None:
            continue
        try:
            hist = yf.Ticker(s["ticker"]).history(period="2d")
            if not hist.empty:
                # Yahoo returns GBp stocks in pence
                s["live_price_p"] = round(float(hist["Close"].iloc[-1]), 2)
        except Exception as e:
            print(f"  Failed to fetch {s['ticker']}: {e}")

    fetched = sum(1 for s in stocks if s["live_price_p"] is not None)
    print(f"  Fetched {fetched}/{len(tickers)} prices.")


def compute_signals(stocks):
    """Compute value ratios and signals from live prices and blended targets."""
    for s in stocks:
        if s["live_price_p"] and s["blended_p"] and s["live_price_p"] > 0:
            s["value_ratio"] = round(s["blended_p"] / s["live_price_p"], 3)
            s["computed_signal"] = compute_signal(s["value_ratio"])
        else:
            s["computed_signal"] = s.get("curr_signal", "N/A")


# ── Markdown generation ──────────────────────────────────────────────────────

def fmt_price(p):
    if p is None:
        return "[PENDING]"
    if p >= 100:
        return f"{p:,.0f}p"
    return f"{p:.1f}p"

def fmt_vr(vr):
    if vr is None:
        return "[PENDING]"
    return f"{vr:.2f}"

def fmt_signal(sig):
    emoji = SIGNAL_EMOJI.get(sig, "")
    if emoji:
        return f"{emoji} {sig}"
    return sig or "N/A"


def generate_markdown(stocks, tracker_date):
    """Generate the full Substack article as markdown."""
    lines = []
    w = lines.append  # shorthand

    # Filter out investment trusts with "No Valuation" and absurd value ratios (data issues)
    valid = [s for s in stocks
             if s["computed_signal"] not in ("N/A", "No Signal", "", None)
             and s.get("model", "") not in ("No Valuation",)
             and (s["value_ratio"] is None or s["value_ratio"] < 10)]  # VR > 10 = data error
    portfolio = [s for s in valid if s["in_portfolio"]]
    non_portfolio = [s for s in valid if not s["in_portfolio"]]

    # Signal changes
    changes = [s for s in valid if s["prev_signal"] and s["curr_signal"]
               and s["prev_signal"] != s["curr_signal"]
               and s["prev_signal"] not in ("No Signal", "N/A", "")]

    # Counts
    strong_buys = [s for s in valid if s["computed_signal"] == "Strong Buy"]
    strong_sells = [s for s in valid if s["computed_signal"] == "Strong Sell"]

    # Group portfolio by signal
    port_by_signal = defaultdict(list)
    for s in portfolio:
        port_by_signal[s["computed_signal"]].append(s)

    # Group non-portfolio strong signals
    np_strong_buy = [s for s in non_portfolio if s["computed_signal"] == "Strong Buy"]
    np_strong_sell = [s for s in non_portfolio if s["computed_signal"] == "Strong Sell"]

    # Sector heatmap
    sector_signals = defaultdict(lambda: defaultdict(int))
    for s in valid:
        sector_signals[s["sector"]][s["computed_signal"]] += 1

    date_str = tracker_date.strftime("%-d %B %Y")
    week_str = tracker_date.strftime("%Y-%m-%d")

    # ── Header ───────────────────────────────────────────────────────────
    w(f"# FTSE Valuation Tracker - Week of {date_str}")
    w("")
    w("*Updated every Tuesday. All valuations from my proprietary DCF/DDM/EPV models.*")
    w("")
    w("---")
    w("")

    # ── Section 1: Summary ───────────────────────────────────────────────
    w("## This Week at a Glance")
    w("")
    w(f"- **{len(changes)} signal change{'s' if len(changes) != 1 else ''}** this week")
    if changes:
        for c in changes:
            w(f"  - {c['company']} ({c['ticker']}): {c['prev_signal']} -> {c['curr_signal']}")
    w(f"- **{len(strong_buys)} Strong Buy** signals across the FTSE universe")
    w(f"- **{len(strong_sells)} Strong Sell** signals - names my models say are overvalued")
    w("")

    # ── Section 2: Signal changes ────────────────────────────────────────
    w("## Signal Changes This Week")
    w("")
    if changes:
        w("*Stocks where the signal moved from last week.*")
        w("")
        w("| Company | Ticker | Sector | Previous | New | Target | Live Price | VR |")
        w("|---|---|---|---|---|---|---|---|")
        for c in sorted(changes, key=lambda x: SIGNAL_ORDER.get(x["computed_signal"], 9)):
            w(f"| {c['company']} | {c['ticker']} | {c['sector']} | "
              f"{fmt_signal(c['prev_signal'])} | {fmt_signal(c['curr_signal'])} | "
              f"{fmt_price(c['blended_p'])} | {fmt_price(c['live_price_p'])} | {fmt_vr(c['value_ratio'])} |")
    else:
        w("No signal changes this week. All valuations stable at current prices.")
    w("")
    w("---")
    w("")

    # ── Section 3: Portfolio table ───────────────────────────────────────
    w("## My Portfolio - Current Signals")
    w("")
    w("*These are the FTSE stocks I hold in my live eToro portfolio. See all positions at "
      "[etoro.com/people/dalkent13](https://www.etoro.com/people/dalkent13).*")
    w("")

    signal_labels = ["Strong Buy", "Buy", "Fair Value", "Sell", "Strong Sell"]
    for sig in signal_labels:
        group = port_by_signal.get(sig, [])
        if not group:
            continue
        w(f"### {fmt_signal(sig)}")
        w("")
        w("| Company | Ticker | Sector | Target | Live Price | VR | Signal |")
        w("|---|---|---|---|---|---|---|")
        for s in sorted(group, key=lambda x: -(x["value_ratio"] or 0)):
            w(f"| {s['company']} | {s['ticker']} | {s['sector']} | "
              f"{fmt_price(s['blended_p'])} | {fmt_price(s['live_price_p'])} | "
              f"{fmt_vr(s['value_ratio'])} | {fmt_signal(s['computed_signal'])} |")
        w("")

    w("---")
    w("")

    # ── Section 4: Beyond portfolio ──────────────────────────────────────
    w("## Beyond My Portfolio - FTSE Strong Signals")
    w("")
    w("*FTSE stocks I don't currently hold where the models are flagging extreme valuations.*")
    w("")

    if np_strong_buy:
        w(f"### {fmt_signal('Strong Buy')} - Not in Portfolio")
        w("")
        w("| Company | Ticker | Sector | Target | Live Price | VR |")
        w("|---|---|---|---|---|---|")
        for s in sorted(np_strong_buy, key=lambda x: -(x["value_ratio"] or 0)):
            w(f"| {s['company']} | {s['ticker']} | {s['sector']} | "
              f"{fmt_price(s['blended_p'])} | {fmt_price(s['live_price_p'])} | {fmt_vr(s['value_ratio'])} |")
        w("")

    if np_strong_sell:
        w(f"### {fmt_signal('Strong Sell')} - Not in Portfolio")
        w("")
        w("| Company | Ticker | Sector | Target | Live Price | VR |")
        w("|---|---|---|---|---|---|")
        for s in sorted(np_strong_sell, key=lambda x: (x["value_ratio"] or 99)):
            w(f"| {s['company']} | {s['ticker']} | {s['sector']} | "
              f"{fmt_price(s['blended_p'])} | {fmt_price(s['live_price_p'])} | {fmt_vr(s['value_ratio'])} |")
        w("")

    w("---")
    w("")

    # ── Section 5: Sector heatmap ────────────────────────────────────────
    w("## Sector Heatmap")
    w("")
    w("| Sector | Strong Buy | Buy | Fair Value | Sell | Strong Sell |")
    w("|---|---|---|---|---|---|")
    for sector in sorted(sector_signals.keys()):
        counts = sector_signals[sector]
        w(f"| {sector} | {counts.get('Strong Buy', 0)} | {counts.get('Buy', 0)} | "
          f"{counts.get('Fair Value', 0)} | {counts.get('Sell', 0)} | {counts.get('Strong Sell', 0)} |")
    w("")

    # Cheapest / most expensive
    sector_buy_pct = {}
    for sector, counts in sector_signals.items():
        total = sum(counts.values())
        buys = counts.get("Strong Buy", 0) + counts.get("Buy", 0)
        if total > 0:
            sector_buy_pct[sector] = buys / total
    cheapest = sorted(sector_buy_pct.items(), key=lambda x: -x[1])[:2]
    most_exp_pct = {}
    for sector, counts in sector_signals.items():
        total = sum(counts.values())
        sells = counts.get("Strong Sell", 0) + counts.get("Sell", 0)
        if total > 0:
            most_exp_pct[sector] = sells / total
    expensive = sorted(most_exp_pct.items(), key=lambda x: -x[1])[:2]

    if cheapest:
        w(f"**Cheapest sectors:** {', '.join(s for s, _ in cheapest)}")
    if expensive:
        w(f"**Most expensive sectors:** {', '.join(s for s, _ in expensive)}")
    w("")
    w("---")
    w("")

    # ── Section 6: Approaching boundary ──────────────────────────────────
    w("## Approaching the Boundary")
    w("")
    w("*Stocks close to a signal change - value ratio within 5% of a threshold.*")
    w("")

    boundaries = [1.25, 1.10, 0.90, 0.75]
    boundary_names = {1.25: "Strong Buy/Buy", 1.10: "Buy/Fair Value",
                      0.90: "Fair Value/Sell", 0.75: "Sell/Strong Sell"}
    near_boundary = []
    for s in valid:
        vr = s["value_ratio"]
        if vr is None:
            continue
        for b in boundaries:
            if abs(vr - b) / b <= 0.05:
                direction = "upgrade" if vr < b else "downgrade"
                near_boundary.append({
                    **s,
                    "boundary": b,
                    "boundary_name": boundary_names[b],
                    "direction": direction,
                })
                break

    if near_boundary:
        w("| Company | Ticker | Signal | VR | Nearest Boundary | Direction |")
        w("|---|---|---|---|---|---|")
        for n in sorted(near_boundary, key=lambda x: abs(x["value_ratio"] - x["boundary"])):
            w(f"| {n['company']} | {n['ticker']} | {fmt_signal(n['computed_signal'])} | "
              f"{fmt_vr(n['value_ratio'])} | {n['boundary_name']} ({n['boundary']:.2f}) | "
              f"Potential {n['direction']} |")
    else:
        w("No stocks currently within 5% of a signal boundary.")
    w("")
    w("---")
    w("")

    # ── Methodology ──────────────────────────────────────────────────────
    w("## Methodology")
    w("")
    w("All valuations use my three-model framework: **DCF** (Discounted Cash Flow), "
      "**DDM** (Dividend Discount Model), and **EPV** (Earnings Power Value). "
      "Models are blended with sector-specific weights. Banks use DDM + P/B Excess Returns "
      "+ EPS Capitalisation (no DCF).")
    w("")
    w("**Current parameters:** UK risk-free rate 4.9% | Equity risk premium 5.0% | Terminal growth 2.5%")
    w("")
    w("Full methodology: [How to Value a Company](https://dalkent13.substack.com/p/how-to-value-a-company)")
    w("")
    w("---")
    w("")
    w("*Not financial advice. These are my personal views based on my own valuation models. "
      "Always do your own research before investing.*")
    w("")
    w("*Neil Daley - CFA Charterholder - "
      "[eToro](https://www.etoro.com/people/dalkent13) - "
      "[X/Twitter](https://x.com/Dalkent13)*")

    return "\n".join(lines)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate FTSE Valuation Tracker for Substack")
    parser.add_argument("--date", type=str, default=None,
                        help="Publication date (YYYY-MM-DD). Defaults to today.")
    parser.add_argument("--no-fetch", action="store_true",
                        help="Skip Yahoo Finance price fetch (use for testing)")
    args = parser.parse_args()

    tracker_date = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today()

    stocks, gbpusd, portfolio_tickers = load_master()
    print(f"Loaded {len(stocks)} FTSE stocks. {len(portfolio_tickers)} in portfolio. GBP/USD: {gbpusd}")

    if not args.no_fetch:
        fetch_prices(stocks)
    else:
        print("Skipping price fetch (--no-fetch).")

    compute_signals(stocks)

    md = generate_markdown(stocks, tracker_date)

    # Write output
    DRAFTS_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"{tracker_date.isoformat()} FTSE Valuation Tracker.md"
    output_path = DRAFTS_DIR / filename
    output_path.write_text(md, encoding="utf-8")
    print(f"\nTracker written to: {output_path}")
    print(f"  Signal changes: {sum(1 for s in stocks if s.get('prev_signal') and s.get('curr_signal') and s['prev_signal'] != s['curr_signal'] and s['prev_signal'] not in ('No Signal', 'N/A', ''))}")
    print(f"  Strong Buys: {sum(1 for s in stocks if s.get('computed_signal') == 'Strong Buy')}")
    print(f"  Strong Sells: {sum(1 for s in stocks if s.get('computed_signal') == 'Strong Sell')}")


if __name__ == "__main__":
    main()
