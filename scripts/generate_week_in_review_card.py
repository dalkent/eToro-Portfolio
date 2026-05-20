"""generate_week_in_review_card.py — render a branded "Week in Review" snapshot
card matching the published Daley Valuations editorial style (cream + serif +
italic accent + dark-navy table header + signal band).

Drives:
  - Top metadata strip:  NEIL DALEY · PHD · CFA   |   FRIDAY · 8 MAY 2026
  - Headline:            "Week in Review" with italic accent on "Review"
  - Italic subtitle:     "FTSE -1.5% · defensives doing the work"
  - 5 KPI cards:         FTSE close · Week move · Strong Buys · Sells · Fair Value
  - Signal band
  - Position table:      6 rows of held tickers with their current signal
  - Footer:              D C F · D D M · E P V   |   DALEYVALUATIONS.SUBSTACK.COM · @DALKENT13

Outputs to:
  Projects\eToro & Investing\Drafts\Article Images\YYYY-MM-DD\
    week-in-review-card.png

Numbers in this file are pinned to the live state on 2026-05-08 read from
combined_portfolio.json and etoro_master.json. Do NOT regenerate values from
memory — re-read the source files if you re-run this for a different week.
"""
from __future__ import annotations

import os
import sys
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib import rcParams

# ── Paths ─────────────────────────────────────────────────────────────────────
# Resolve VAULT via paths.py so this works on Windows and macOS. VAULT_ROOT env
# var still wins if set explicitly.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import VAULT_DIR  # noqa: E402

VAULT = Path(os.environ.get("VAULT_ROOT", str(VAULT_DIR)))
OUT_DIR = VAULT / "Projects" / "eToro & Investing" / "Drafts" / "Article Images" / "2026-05-08"
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PATH = OUT_DIR / "week-in-review-card.png"

# ── Editorial palette (matches the Substack live tracker hero) ────────────────
BG_CREAM   = "#f0ece1"
INK_DARK   = "#1d3324"   # forest green for headlines
INK_BODY   = "#2a3933"
INK_MUTED  = "#7a7c75"
INK_RULE   = "#8a8a82"
CARD_BORDER = "#7a7c75"
HEADER_BG  = "#1e293b"
HEADER_FG  = "white"
ROW_TINT   = "#faf7ee"

SIG_COLOURS = {
    "Strong Buy":  "#2d5a40",
    "Buy":         "#5a8a5e",
    "Fair Value":  "#c5a23a",
    "Sell":        "#b8693d",
    "Strong Sell": "#8e3a2a",
}
SIG_BAR = {
    "Strong Buy":  "#3a6e4f",
    "Buy":         "#6ea177",
    "Fair Value":  "#d4b04c",
    "Sell":        "#c87a4f",
    "Strong Sell": "#9e4536",
}

# Serif fonts available on Windows
rcParams["font.family"] = "serif"
rcParams["font.serif"] = ["Georgia", "Cambria", "Times New Roman", "DejaVu Serif"]
rcParams["text.color"] = INK_DARK


def spaced(s: str, gap: str = "  ") -> str:
    return gap.join(list(s))


# ── Content for this week (8 May 2026) ───────────────────────────────────────
DATE_LABEL = "FRIDAY  ·  8  MAY  2026"

# KPI cards — value, label, colour
KPI_CARDS = [
    ("10,213",  "FTSE 100 CLOSE",   INK_DARK),
    ("-1.5%",   "WEEK MOVE",        SIG_COLOURS["Sell"]),
    ("3",       "STRONG BUYS HELD", SIG_COLOURS["Strong Buy"]),
    ("2",       "SELLS HELD",       SIG_COLOURS["Sell"]),
    ("0",       "TRADES THIS WEEK", INK_BODY),
]

# Snapshot table rows (Ticker, Company, Signal, ROI %)
TABLE_ROWS = [
    ("$BATS.L",  "British American Tobacco",  "Strong Buy",  "+50%"),
    ("$BP.L",    "BP",                         "Strong Buy",  "+38%"),
    ("$IMB.L",   "Imperial Brands",            "Buy",         "+32%"),
    ("$NWG.L",   "NatWest Group",              "Fair Value",  "+3%"),
    ("$VOD.L",   "Vodafone",                   "Sell",        "+53%"),
    ("$TATE.L",  "Tate & Lyle",                "Sell",        "-47%"),
]


# ── Drawing helpers ──────────────────────────────────────────────────────────
def draw_card(ax, x, y, w, h, *, label, value, value_color):
    """Draw a single KPI card with a tracking-wide top label and big number."""
    rect = patches.FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.0,rounding_size=0.05",
        linewidth=0.9, edgecolor=CARD_BORDER, facecolor=BG_CREAM,
    )
    ax.add_patch(rect)
    # Tracking-wide label centred near top of the card
    ax.text(x + w / 2, y + h - 0.18,
            spaced(label, "  "),
            ha="center", va="top",
            fontsize=8.0, color=INK_MUTED, family="sans-serif",
            fontweight="bold")
    # Big value centred
    ax.text(x + w / 2, y + h * 0.42,
            value,
            ha="center", va="center",
            fontsize=28, color=value_color, family="serif")


def draw_signal_band(ax, x0, y0, w, h):
    """5-segment signal band: Strong Buy → Strong Sell."""
    order = ["Strong Buy", "Buy", "Fair Value", "Sell", "Strong Sell"]
    seg_w = w / len(order)
    for i, s in enumerate(order):
        rect = patches.Rectangle(
            (x0 + i * seg_w, y0), seg_w, h,
            facecolor=SIG_BAR[s], edgecolor="none",
        )
        ax.add_patch(rect)
        ax.text(x0 + i * seg_w + seg_w / 2, y0 + h / 2,
                s.upper(),
                ha="center", va="center",
                fontsize=8.5, color="white",
                family="sans-serif", fontweight="bold")


def draw_position_table(ax, x0, y0, w, h, rows):
    """Render a 4-column data table: Ticker | Company | Signal | ROI."""
    n_rows = len(rows)
    # Header band
    header_h = 0.55
    body_h = h - header_h
    row_h = body_h / n_rows

    # Column proportions (sum = 1.0)
    cols = [0.18, 0.45, 0.22, 0.15]
    col_x = [x0]
    for c in cols[:-1]:
        col_x.append(col_x[-1] + c * w)

    # Header background
    header_rect = patches.Rectangle(
        (x0, y0 + body_h), w, header_h,
        facecolor=HEADER_BG, edgecolor="none",
    )
    ax.add_patch(header_rect)
    headers = ["TICKER", "COMPANY", "SIGNAL", "ROI"]
    for i, hd in enumerate(headers):
        ha = "left" if i < 2 else "right" if i == 3 else "left"
        x_text = col_x[i] + 0.18 if ha == "left" else col_x[i] + cols[i] * w - 0.18
        ax.text(x_text, y0 + body_h + header_h / 2,
                spaced(hd, " "),
                ha=ha, va="center",
                fontsize=9.5, color=HEADER_FG,
                family="sans-serif", fontweight="bold")

    # Body rows
    for r_idx, row in enumerate(rows):
        ry = y0 + body_h - (r_idx + 1) * row_h
        if r_idx % 2 == 1:
            tint = patches.Rectangle(
                (x0, ry), w, row_h,
                facecolor=ROW_TINT, edgecolor="none",
            )
            ax.add_patch(tint)
        ticker, company, signal, roi = row
        sig_colour = SIG_COLOURS.get(signal, INK_BODY)
        roi_colour = "#a8332e" if roi.startswith("-") else SIG_COLOURS["Strong Buy"]

        # Ticker (serif, bold-feel via size)
        ax.text(col_x[0] + 0.18, ry + row_h / 2, ticker,
                ha="left", va="center",
                fontsize=12.5, color=INK_DARK, family="serif")
        # Company
        ax.text(col_x[1] + 0.18, ry + row_h / 2, company,
                ha="left", va="center",
                fontsize=12, color=INK_BODY, family="serif")
        # Signal — coloured text only, no pill
        ax.text(col_x[2] + 0.18, ry + row_h / 2, signal,
                ha="left", va="center",
                fontsize=12, color=sig_colour, family="serif",
                fontweight="bold")
        # ROI right-aligned
        ax.text(col_x[3] + cols[3] * w - 0.18, ry + row_h / 2, roi,
                ha="right", va="center",
                fontsize=12.5, color=roi_colour, family="serif")

        # Thin row separator
        ax.plot([x0, x0 + w], [ry, ry],
                color="#e6e0d0", linewidth=0.7)


# ── Figure ────────────────────────────────────────────────────────────────────
def main():
    FIG_W, FIG_H = 16, 10
    fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=200)
    fig.patch.set_facecolor(BG_CREAM)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, FIG_W)
    ax.set_ylim(0, FIG_H)
    ax.set_facecolor(BG_CREAM)
    ax.set_axis_off()

    # ── Top metadata strip ────────────────────────────────────────────────────
    top_y = FIG_H - 0.55
    ax.text(0.55, top_y, spaced("NEIL DALEY · PHD · CFA", " "),
            ha="left", va="center",
            fontsize=9.0, color=INK_MUTED,
            family="sans-serif", fontweight="bold")
    ax.text(FIG_W - 0.55, top_y, spaced(DATE_LABEL, " "),
            ha="right", va="center",
            fontsize=9.0, color=INK_MUTED,
            family="sans-serif", fontweight="bold")
    ax.plot([0.55, FIG_W - 0.55], [top_y - 0.25, top_y - 0.25],
            color=INK_RULE, linewidth=0.7, alpha=0.65)

    # ── Headline + subtitle ───────────────────────────────────────────────────
    head_y = FIG_H - 1.50
    # Build "Week in Review" with italic accent on "Review"
    ax.text(0.55, head_y, "Week in ",
            ha="left", va="center",
            fontsize=58, color=INK_DARK, family="serif")
    # Estimate width of "Week in " to position the italic part
    # (matplotlib doesn't give us exact metrics easily; we offset visually)
    ax.text(4.30, head_y, "Review",
            ha="left", va="center",
            fontsize=58, color=INK_DARK, family="serif",
            fontstyle="italic")

    sub_y = head_y - 1.05
    ax.text(0.55, sub_y,
            "FTSE 100 -1.5% on the week  ·  defensives doing the work  ·  one Sell signal worth flagging",
            ha="left", va="center",
            fontsize=15, color=INK_MUTED, family="serif",
            fontstyle="italic")

    # ── KPI cards (5 across) ──────────────────────────────────────────────────
    kpi_y = sub_y - 1.85
    kpi_h = 1.30
    margin = 0.55
    gap = 0.20
    n = len(KPI_CARDS)
    avail = FIG_W - 2 * margin - (n - 1) * gap
    card_w = avail / n
    for i, (val, lbl, col) in enumerate(KPI_CARDS):
        x = margin + i * (card_w + gap)
        draw_card(ax, x, kpi_y, card_w, kpi_h,
                  label=lbl, value=val, value_color=col)

    # ── Signal band ───────────────────────────────────────────────────────────
    band_y = kpi_y - 0.55
    draw_signal_band(ax, margin, band_y, FIG_W - 2 * margin, 0.32)

    # ── Section label above table ─────────────────────────────────────────────
    sec_y = band_y - 0.55
    ax.text(margin, sec_y,
            spaced("SIX HOLDINGS · WHAT THE MODEL SAYS NOW", " "),
            ha="left", va="center",
            fontsize=10, color=INK_MUTED,
            family="sans-serif", fontweight="bold")
    ax.plot([margin, FIG_W - margin], [sec_y - 0.18, sec_y - 0.18],
            color=INK_RULE, linewidth=0.6, alpha=0.55)

    # ── Position table ────────────────────────────────────────────────────────
    table_h = 2.55
    table_y = sec_y - 0.30 - table_h
    draw_position_table(ax, margin, table_y, FIG_W - 2 * margin, table_h, TABLE_ROWS)

    # ── Footer ────────────────────────────────────────────────────────────────
    foot_y = 0.40
    ax.plot([margin, FIG_W - margin], [foot_y + 0.30, foot_y + 0.30],
            color=INK_RULE, linewidth=0.6, alpha=0.55)
    ax.text(margin, foot_y, spaced("D C F  ·  D D M  ·  E P V", " "),
            ha="left", va="center",
            fontsize=9.0, color=INK_MUTED,
            family="sans-serif", fontweight="bold")
    ax.text(FIG_W - margin, foot_y,
            spaced("DALEYVALUATIONS.SUBSTACK.COM  ·  @DALKENT13", " "),
            ha="right", va="center",
            fontsize=9.0, color=INK_MUTED,
            family="sans-serif", fontweight="bold")

    fig.savefig(OUT_PATH, dpi=200, facecolor=BG_CREAM)
    plt.close(fig)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
