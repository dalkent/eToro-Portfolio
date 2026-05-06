"""generate_tracker_hero.py — render the cream editorial-magazine hero card
for the Tuesday Substack article.

Matches the visual language of the live published tracker articles:
  https://daleyvaluations.substack.com/p/ftse-valuation-tracker-week-of-28

Layout:
  - Top rule with NEIL DALEY · PHD · CFA  /  WEEK OF · DD MONTH YYYY
  - Big serif headline "FTSE Valuation Tracker" (italic "Valuation")
  - Italic subtitle "Institutional-grade valuations · NN FTSE companies · weekly"
  - 5 signal-count cards in a row (thin borders, coloured numbers)
  - Color-graded horizontal signal bar
  - "TOP 6 PORTFOLIO STRONG BUYS · RANKED BY VALUE RATIO" sub-rule
  - 3 × 2 grid of stock cards (sector, company, ticker, VR)
  - Footer: DCF · DDM · EPV  /  DALKENT13.SUBSTACK.COM · @DALKENT13

Uses the same site loader as generate_tracker.py / generate_tracker_images.py
so signal counts and DVRs match daleyvaluations.com exactly.

Output:
  Drafts\\Article Images\\YYYY-MM-DD tracker_tables\\00_hero.png

Usage:
  python scripts/generate_tracker_hero.py
  python scripts/generate_tracker_hero.py --date 2026-05-05
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib import rcParams

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent.parent
SITE_REPO = Path(r"C:\Users\Neil\ClaudeCode\daleyvaluations-site")
VAULT     = Path(os.environ.get("VAULT_ROOT", r"C:\Users\Neil\My Drive\Daley's Brain"))
DRAFTS    = VAULT / "Projects" / "eToro & Investing" / "Drafts"

# ── Editorial colour palette (matches the Substack live article) ──────────────
BG_CREAM     = "#f0ece1"   # warm cream background
INK_DARK     = "#1d3324"   # dark forest green for headlines
INK_BODY     = "#2a3933"   # body text
INK_MUTED    = "#7a7c75"   # muted gray for labels / meta
INK_RULE     = "#8a8a82"   # rule lines
CARD_BORDER  = "#7a7c75"   # thin card borders

# Signal colours — matching the published palette
SIG_COLOURS = {
    "Strong Buy":  "#2d5a40",   # dark green
    "Buy":         "#5a8a5e",   # medium green
    "Fair Value":  "#c5a23a",   # mustard yellow
    "Sell":        "#b8693d",   # burnt orange
    "Strong Sell": "#8e3a2a",   # dark red
}
SIG_BAR_LIGHT = {
    "Strong Buy":  "#3a6e4f",
    "Buy":         "#6ea177",
    "Fair Value":  "#d4b04c",
    "Sell":        "#c87a4f",
    "Strong Sell": "#9e4536",
}

# Serif font setup — try a serif that's available on Windows
# (DejaVu Serif is matplotlib's bundled fallback; Cambria/Georgia exist on Windows)
rcParams["font.family"] = "serif"
rcParams["font.serif"] = ["Georgia", "Cambria", "Times New Roman", "DejaVu Serif"]
rcParams["text.color"] = INK_DARK


# ── Data load (mirror of generate_tracker_images.py) ──────────────────────────
def load_records():
    sys.path.insert(0, str(SITE_REPO / "scripts"))
    import build_site  # type: ignore
    import importlib; importlib.reload(build_site)
    data = build_site.load_data(build_site.DEFAULT_DATA_FILE)
    held = build_site.load_held_tickers(build_site.DEFAULT_PORTFOLIO_FILE)
    all_recs = build_site.join_records(data)
    public = build_site.filter_public(all_recs)
    prices = build_site.fetch_live_prices(public, force_refresh=False)
    public = build_site.apply_live_prices(public, prices)
    for r in public:
        sig, _ = build_site.signal_for(r.get("value_ratio"))
        r["computed_signal"] = sig
    return public, held


# ── Hero rendering ────────────────────────────────────────────────────────────
def render_hero(records, held, output_path, hero_date: date):
    held_upper = {h.upper() for h in (held or set())}
    in_port = lambda r: (r.get("yahoo_ticker") or r["ticker"]).upper() in held_upper

    # 1. Signal counts
    counts = Counter()
    for r in records:
        counts[r.get("computed_signal") or "N/A"] += 1
    sig_order = ["Strong Buy", "Buy", "Fair Value", "Sell", "Strong Sell"]
    sig_vals  = [counts.get(s, 0) for s in sig_order]
    universe  = sum(sig_vals)

    # 2. Top 6 portfolio Strong Buys by VR
    top6 = sorted(
        [r for r in records if in_port(r) and r.get("computed_signal") == "Strong Buy"],
        key=lambda r: -(r.get("value_ratio") or 0),
    )[:6]

    # 3. Layout — 14 × 9 inches at 100 dpi → 1400 × 900 px
    fig = plt.figure(figsize=(14, 9), dpi=100, facecolor=BG_CREAM)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.set_facecolor(BG_CREAM)

    # ── Top rule + author / date ─────────────────────────────────────────────
    week_label = hero_date.strftime("%-d %B %Y").upper() if os.name != "nt" else \
                 hero_date.strftime("%#d %B %Y").upper()
    # Top rule line
    ax.add_line(plt.Line2D([0.05, 0.95], [0.945, 0.945],
                            color=INK_DARK, linewidth=0.8, transform=ax.transAxes))
    ax.text(0.05, 0.962, "NEIL DALEY  ·  PHD  ·  CFA",
            fontsize=9, color=INK_MUTED, ha="left", va="bottom",
            family="sans-serif", weight="normal", transform=ax.transAxes)
    ax.text(0.95, 0.962, f"WEEK OF  ·  {week_label}",
            fontsize=9, color=INK_MUTED, ha="right", va="bottom",
            family="sans-serif", weight="normal", transform=ax.transAxes)

    # ── Title (FTSE Valuation Tracker — italic on "Valuation") ───────────────
    # Three text() calls so we can italicise just the middle word
    title_y = 0.78
    title_size = 64
    # Render full string measurement-style: place start, then advance by widths
    # We approximate widths by rendering individual text objects centred relative.
    # Simpler approach: anchor each piece with explicit x positions tuned to look right.
    ax.text(0.05, title_y, "FTSE", fontsize=title_size, color=INK_DARK,
            ha="left", va="baseline", family="serif", weight="normal",
            transform=ax.transAxes)
    ax.text(0.225, title_y, "Valuation", fontsize=title_size, color=INK_DARK,
            ha="left", va="baseline", family="serif", style="italic",
            weight="normal", transform=ax.transAxes)
    ax.text(0.585, title_y, "Tracker", fontsize=title_size, color=INK_DARK,
            ha="left", va="baseline", family="serif", weight="normal",
            transform=ax.transAxes)

    # Italic subtitle
    ax.text(0.05, 0.715,
            f"Institutional-grade valuations  ·  {universe} FTSE companies  ·  weekly",
            fontsize=15, color=INK_MUTED, ha="left", va="top",
            family="serif", style="italic", transform=ax.transAxes)

    # ── 5 signal cards row ───────────────────────────────────────────────────
    card_y      = 0.51
    card_h      = 0.13
    card_pad    = 0.012
    card_total_w = 0.90
    card_w      = (card_total_w - 4 * card_pad) / 5
    for i, sig in enumerate(sig_order):
        x = 0.05 + i * (card_w + card_pad)
        # Card border
        ax.add_patch(patches.Rectangle((x, card_y), card_w, card_h,
                                        transform=ax.transAxes,
                                        facecolor="none", edgecolor=INK_DARK,
                                        linewidth=0.6, zorder=1))
        # Label (top left)
        label_text = sig.upper()
        ax.text(x + 0.012, card_y + card_h - 0.018, label_text,
                fontsize=8, color=INK_MUTED, ha="left", va="top",
                family="sans-serif", weight="normal", transform=ax.transAxes,
                bbox=dict(facecolor=BG_CREAM, edgecolor="none", pad=2))
        # Big number (right side, dominant)
        num = sig_vals[i]
        ax.text(x + card_w - 0.020, card_y + 0.015, str(num),
                fontsize=42, color=SIG_COLOURS[sig], ha="right", va="bottom",
                family="serif", weight="normal", transform=ax.transAxes)

    # ── Signal bar (proportional segments) ───────────────────────────────────
    bar_y = card_y - 0.038
    bar_h = 0.012
    if universe > 0:
        cur_x = 0.05
        for sig in sig_order:
            n = sig_vals[sig_order.index(sig)]
            if n == 0:
                continue
            seg_w = (n / universe) * card_total_w
            ax.add_patch(patches.Rectangle((cur_x, bar_y), seg_w, bar_h,
                                            transform=ax.transAxes,
                                            facecolor=SIG_BAR_LIGHT[sig],
                                            edgecolor="none", zorder=1))
            cur_x += seg_w

    # ── Top 6 sub-rule ───────────────────────────────────────────────────────
    sub_y = bar_y - 0.04
    # short black line then text
    ax.add_line(plt.Line2D([0.05, 0.085], [sub_y, sub_y],
                            color=INK_DARK, linewidth=1.0, transform=ax.transAxes))
    ax.text(0.10, sub_y, "TOP 6 PORTFOLIO STRONG BUYS  ·  RANKED BY VALUE RATIO",
            fontsize=9.5, color=INK_DARK, ha="left", va="center",
            family="sans-serif", weight="bold",
            transform=ax.transAxes)

    # ── Top picks 3×2 grid ───────────────────────────────────────────────────
    pick_y_start = sub_y - 0.03
    pick_h       = 0.115
    pick_pad     = 0.012
    pick_total_w = 0.90
    pick_w       = (pick_total_w - 2 * pick_pad) / 3

    for i, r in enumerate(top6):
        col = i % 3
        row = i // 3
        x = 0.05 + col * (pick_w + pick_pad)
        y = pick_y_start - pick_h - (row * (pick_h + pick_pad))
        # Card border
        ax.add_patch(patches.Rectangle((x, y), pick_w, pick_h,
                                        transform=ax.transAxes,
                                        facecolor="none", edgecolor=INK_DARK,
                                        linewidth=0.6, zorder=1))
        # Sector label (top-left, small caps)
        sector = (r.get("sector") or "").upper()
        ax.text(x + 0.012, y + pick_h - 0.018, sector,
                fontsize=8, color=INK_MUTED, ha="left", va="top",
                family="sans-serif", weight="normal",
                bbox=dict(facecolor=BG_CREAM, edgecolor="none", pad=2),
                transform=ax.transAxes)
        # VALUE RATIO label (top-right)
        ax.text(x + pick_w - 0.012, y + pick_h - 0.018, "VALUE RATIO",
                fontsize=8, color=INK_MUTED, ha="right", va="top",
                family="sans-serif", weight="normal",
                bbox=dict(facecolor=BG_CREAM, edgecolor="none", pad=2),
                transform=ax.transAxes)
        # Company name (left, big serif)
        company = r.get("company") or r["ticker"]
        # Truncate if too long
        if len(company) > 28:
            company = company[:26] + "…"
        ax.text(x + 0.012, y + pick_h * 0.42, company,
                fontsize=18, color=INK_DARK, ha="left", va="center",
                family="serif", weight="normal",
                transform=ax.transAxes)
        # Ticker (left, small mono)
        ax.text(x + 0.012, y + 0.012, r["ticker"],
                fontsize=10, color=INK_MUTED, ha="left", va="bottom",
                family="monospace", weight="normal",
                transform=ax.transAxes)
        # Big VR (right)
        vr = r.get("value_ratio")
        vr_text = f"{vr:.2f}" if vr is not None else "—"
        ax.text(x + pick_w - 0.012, y + pick_h * 0.30, vr_text,
                fontsize=42, color=SIG_COLOURS["Strong Buy"], ha="right", va="center",
                family="serif", weight="normal",
                transform=ax.transAxes)

    # ── Footer ───────────────────────────────────────────────────────────────
    ax.add_line(plt.Line2D([0.05, 0.95], [0.060, 0.060],
                            color=INK_DARK, linewidth=0.8, transform=ax.transAxes))
    ax.text(0.05, 0.038, "DCF  ·  DDM  ·  EPV",
            fontsize=9, color=INK_MUTED, ha="left", va="top",
            family="sans-serif", weight="normal", transform=ax.transAxes)
    ax.text(0.95, 0.038, "DALEYVALUATIONS.SUBSTACK.COM  ·  @DALKENT13",
            fontsize=9, color=INK_MUTED, ha="right", va="top",
            family="sans-serif", weight="normal", transform=ax.transAxes)

    fig.savefig(output_path, dpi=120, facecolor=BG_CREAM, bbox_inches=None,
                pad_inches=0)
    plt.close(fig)
    print(f"  Saved hero: {output_path}")


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", type=str, default=None,
                        help="Date stamp for output folder + 'WEEK OF' label (YYYY-MM-DD)")
    parser.add_argument("--out-dir", type=str, default=None,
                        help="Override output directory")
    args = parser.parse_args()

    hero_date = (datetime.strptime(args.date, "%Y-%m-%d").date()
                 if args.date else date.today())
    stamp = hero_date.isoformat()
    out_dir = Path(args.out_dir) if args.out_dir else (
        DRAFTS / "Article Images" / f"{stamp} tracker_tables"
    )
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output dir: {out_dir}")

    print("Loading records via site loader (matches website universe) …")
    records, held = load_records()
    print(f"  {len(records)} publishable FTSE equities, {len(held or set())} held.")

    output_path = out_dir / "00_hero.png"
    render_hero(records, held, str(output_path), hero_date)
    print()
    print(f"Done. Hero saved to: {output_path}")


if __name__ == "__main__":
    main()
