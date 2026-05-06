"""generate_tracker_images.py — render the 6 PNG tracker tables for Substack.

Reads via daleyvaluations-site/scripts/build_site.py so the universe + filters
+ prices match the live website AND the tracker .md exactly. Writes 6 PNGs to:

  C:\\Users\\Neil\\My Drive\\Daley's Brain\\Projects\\eToro & Investing\\Drafts\\
    Article Images\\YYYY-MM-DD tracker_tables\\
      01_portfolio_strong_buy.png
      02_portfolio_remaining.png      (Buy + Fair Value + Sell + Strong Sell)
      03_beyond_strong_buy.png        (FTSE strong buys not held)
      04_beyond_strong_sell.png       (FTSE strong sells not held)
      05_sector_heatmap.png
      06_signal_changes.png           (only if there are changes this week)

Designed to run as part of the Monday 5pm "lock-in" task. Pure-matplotlib so
it works in any local Python with matplotlib + Pillow available — no
playwright, no chromium, no external HTML→PNG tooling needed.

Usage:
  python scripts/generate_tracker_images.py
  python scripts/generate_tracker_images.py --date 2026-05-05  (override date)
  python scripts/generate_tracker_images.py --out-dir <path>   (override out path)
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # headless / file-only
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent.parent
SITE_REPO = Path(r"C:\Users\Neil\ClaudeCode\daleyvaluations-site")
VAULT     = Path(os.environ.get("VAULT_ROOT", r"C:\Users\Neil\My Drive\Daley's Brain"))
DRAFTS    = VAULT / "Projects" / "eToro & Investing" / "Drafts"

# ── Visual style — matches the playwright-generated tables in archive/unused ──
SIGNAL_BG = {
    "Strong Buy":  "#f0fdf4",
    "Buy":         "#f7fef9",
    "Fair Value":  "#fefce8",
    "Sell":        "#fff7ed",
    "Strong Sell": "#fef2f2",
}
SIGNAL_LABEL = {
    # Plain text — matplotlib's default font can't render colour emojis.
    # Row tinting (SIGNAL_BG) already conveys the signal visually.
    "Strong Buy":  "Strong Buy",
    "Buy":         "Buy",
    "Fair Value":  "Fair Value",
    "Sell":        "Sell",
    "Strong Sell": "Strong Sell",
}
HEADER_BG  = "#1e293b"
HEADER_FG  = "white"
ROW_ALT    = "#f8fafc"
BORDER     = "#e2e8f0"
HEATMAP_FG = {
    1: "#16a34a", 2: "#22c55e", 3: "#ca8a04", 4: "#ea580c", 5: "#dc2626",
}
HEATMAP_BG = {
    1: "#f0fdf4", 2: "#f7fef9", 3: "#fefce8", 4: "#fff7ed", 5: "#fef2f2",
}


# ── Data loading via the site's loader (guarantees consistency) ──────────────
def load_records():
    """Returns (records, held_tickers) using the site's load+filter+price chain."""
    sys.path.insert(0, str(SITE_REPO / "scripts"))
    import build_site  # type: ignore
    import importlib; importlib.reload(build_site)

    data = build_site.load_data(build_site.DEFAULT_DATA_FILE)
    held = build_site.load_held_tickers(build_site.DEFAULT_PORTFOLIO_FILE)
    all_recs = build_site.join_records(data)
    public = build_site.filter_public(all_recs)
    prices = build_site.fetch_live_prices(public, force_refresh=False)
    public = build_site.apply_live_prices(public, prices)
    # signal labels
    for r in public:
        sig, _ = build_site.signal_for(r.get("value_ratio"))
        r["computed_signal"] = sig
    return public, held


# ── Formatting helpers ───────────────────────────────────────────────────────
def fmt_target_pence(v):
    if v is None: return "—"
    return f"{v * 100:,.1f}p" if v < 50 else f"{v * 100:,.0f}p"

def fmt_price_pence(v):
    if v is None: return "—"
    return f"{v:,.1f}p" if v < 50 else f"{v:,.0f}p"

def fmt_vr(v):
    if v is None: return "—"
    return f"{v:.2f}"


# ── Table rendering ──────────────────────────────────────────────────────────
DPI = 120


def _make_axes_filling_figure(fig_w_in, fig_h_in):
    """Create a figure whose axes fill the entire figure with NO subplot margins.

    Using fig.add_axes([0, 0, 1, 1]) is essential — plt.subplots() leaves
    ~12.5% margins on each side by default, which causes table column
    coordinates (specified in inches) to misalign with the visible drawing
    area, leading to text overflowing into adjacent columns.
    """
    fig = plt.figure(figsize=(fig_w_in, fig_h_in), dpi=DPI)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    ax.set_xlim(0, fig_w_in)
    ax.set_ylim(0, fig_h_in)
    return fig, ax


def _measure_text_width_px(text, fontsize, ax, renderer):
    """Return rendered width of `text` in display pixels at the given fontsize."""
    t = ax.text(0, 0, str(text), fontsize=fontsize)
    bb = t.get_window_extent(renderer=renderer)
    w = bb.width
    t.remove()
    return w


def render_table(headers, rows, row_signals, output_path,
                 font_size=11, header_font_size=11,
                 row_height_in=0.42, padding_in=0.12,
                 inner_pad_px=16, min_col_px=60,
                 col_widths=None, title=None):
    """Generic table renderer with auto-sized columns.

    Each column is sized to fit its widest cell content at the chosen font.
    No fixed figure width — the figure expands to fit all data.

    headers: list of column header strings.
    rows: list of [str cells].
    row_signals: parallel list of signal label strings for row tinting.
    col_widths: IGNORED (kept for backward compat with existing call sites).
    """
    n_rows = len(rows)
    n_cols = len(headers)

    # Boot a tiny figure to get a renderer for measuring text widths
    boot_fig, boot_ax = _make_axes_filling_figure(2, 2)
    boot_fig.canvas.draw()
    brenderer = boot_fig.canvas.get_renderer()
    col_widths_px = []
    for ci in range(n_cols):
        w = _measure_text_width_px(headers[ci], header_font_size,
                                   boot_ax, brenderer)
        for row in rows:
            if ci < len(row):
                w = max(w, _measure_text_width_px(row[ci], font_size,
                                                  boot_ax, brenderer))
        col_widths_px.append(max(min_col_px, w + inner_pad_px * 2))
    plt.close(boot_fig)

    fig_w_in = sum(col_widths_px) / DPI + padding_in * 2
    fig_h_in = (n_rows + 1) * row_height_in + padding_in * 2
    if title:
        fig_h_in += 0.4
    fig, ax = _make_axes_filling_figure(fig_w_in, fig_h_in)

    # Column x boundaries in inches
    xs_in = [padding_in]
    for w_px in col_widths_px:
        xs_in.append(xs_in[-1] + w_px / DPI)

    y_top = fig_h_in - padding_in
    if title:
        ax.text(fig_w_in / 2, y_top - 0.05, title,
                ha="center", va="top", fontsize=12, fontweight="bold")
        y_top -= 0.4

    # Header row
    y_h_b = y_top - row_height_in
    ax.add_patch(patches.Rectangle((xs_in[0], y_h_b),
                                    xs_in[-1] - xs_in[0], row_height_in,
                                    facecolor=HEADER_BG, edgecolor="none",
                                    zorder=0))
    for ci, h in enumerate(headers):
        cx = (xs_in[ci] + xs_in[ci + 1]) / 2
        ax.text(cx, y_h_b + row_height_in / 2, h,
                ha="center", va="center",
                fontsize=header_font_size, color=HEADER_FG,
                fontweight="bold")

    # Data rows
    inner_in = inner_pad_px / DPI
    y = y_h_b
    for ri, row in enumerate(rows):
        sig = row_signals[ri] if ri < len(row_signals) else None
        bg = SIGNAL_BG.get(sig, ROW_ALT if ri % 2 == 0 else "#ffffff")
        y -= row_height_in
        ax.add_patch(patches.Rectangle((xs_in[0], y),
                                        xs_in[-1] - xs_in[0], row_height_in,
                                        facecolor=bg, edgecolor=BORDER,
                                        linewidth=0.5, zorder=0))
        for ci, val in enumerate(row):
            ha = "left" if ci < 3 else "right"
            tx = xs_in[ci] + inner_in if ha == "left" else xs_in[ci + 1] - inner_in
            ax.text(tx, y + row_height_in / 2, str(val),
                    ha=ha, va="center", fontsize=font_size,
                    color="#0f172a")

    fig.savefig(output_path, dpi=DPI, facecolor="white",
                bbox_inches=None, pad_inches=0)
    plt.close(fig)
    print(f"  Saved: {output_path}")


def render_heatmap(rows, output_path, title="Sector Heatmap", font_size=11):
    """rows: list of (sector, [sb, b, fv, s, ss]) tuples.

    Auto-sizes the Sector column to fit the widest sector name and the
    five count columns to fit their headers, so nothing gets clipped.
    """
    sigs = ["Strong Buy", "Buy", "Fair Value", "Sell", "Strong Sell"]
    headers = ["Sector"] + sigs
    if not rows:
        return

    boot_fig, boot_ax = _make_axes_filling_figure(2, 2)
    boot_fig.canvas.draw()
    brenderer = boot_fig.canvas.get_renderer()
    pad_px = 18
    sec_w = max(_measure_text_width_px(sector, font_size, boot_ax, brenderer)
                for sector, _ in rows) + pad_px * 2
    sec_w = max(sec_w,
                _measure_text_width_px("Sector", font_size, boot_ax, brenderer)
                + pad_px * 2)
    cnt_w = max(_measure_text_width_px(h, font_size, boot_ax, brenderer)
                for h in sigs) + pad_px * 2
    plt.close(boot_fig)
    col_widths_px = [sec_w] + [cnt_w] * 5

    pad_in = 0.12
    rh = 0.42
    fig_w_in = sum(col_widths_px) / DPI + pad_in * 2
    fig_h_in = (len(rows) + 1) * rh + pad_in * 2
    fig, ax = _make_axes_filling_figure(fig_w_in, fig_h_in)

    xs_in = [pad_in]
    for w in col_widths_px:
        xs_in.append(xs_in[-1] + w / DPI)

    y_top = fig_h_in - pad_in
    y_h_b = y_top - rh
    ax.add_patch(patches.Rectangle((xs_in[0], y_h_b),
                                    xs_in[-1] - xs_in[0], rh,
                                    facecolor=HEADER_BG, edgecolor="none", zorder=0))
    for ci, h in enumerate(headers):
        cx = (xs_in[ci] + xs_in[ci + 1]) / 2
        ax.text(cx, y_h_b + rh / 2, h, ha="center", va="center",
                fontsize=font_size, color=HEADER_FG, fontweight="bold")

    y = y_h_b
    for ri, (sector, counts) in enumerate(rows):
        y -= rh
        ax.add_patch(patches.Rectangle((xs_in[0], y),
                                        xs_in[1] - xs_in[0], rh,
                                        facecolor=ROW_ALT if ri % 2 else "#ffffff",
                                        edgecolor=BORDER, linewidth=0.5, zorder=0))
        ax.text(xs_in[0] + 0.10, y + rh / 2, sector,
                ha="left", va="center", fontsize=font_size, color="#0f172a")
        for ci, val in enumerate(counts, start=1):
            if val == 0:
                bg, fg, weight = "#ffffff", "#cbd5e1", "normal"
            else:
                bg = HEATMAP_BG.get(ci, "#fff")
                fg = HEATMAP_FG.get(ci, "#000")
                weight = "bold"
            ax.add_patch(patches.Rectangle((xs_in[ci], y),
                                            xs_in[ci + 1] - xs_in[ci], rh,
                                            facecolor=bg, edgecolor=BORDER,
                                            linewidth=0.5, zorder=0))
            cx = (xs_in[ci] + xs_in[ci + 1]) / 2
            ax.text(cx, y + rh / 2, str(val), ha="center", va="center",
                    fontsize=12, color=fg, fontweight=weight)

    fig.savefig(output_path, dpi=DPI, facecolor="white",
                bbox_inches=None, pad_inches=0)
    plt.close(fig)
    print(f"  Saved: {output_path}")


# ── Build the 6 image groups from records ────────────────────────────────────
def build_groups(records, held):
    """Returns dict of group_name -> rows."""
    held_upper = {h.upper() for h in (held or set())}
    in_port = lambda r: (r.get("yahoo_ticker") or r["ticker"]).upper() in held_upper

    portfolio = [r for r in records if in_port(r)]
    beyond    = [r for r in records if not in_port(r)]

    def _signal(r):
        return r.get("computed_signal") or "N/A"

    return {
        "portfolio_sb": [r for r in portfolio if _signal(r) == "Strong Buy"],
        "portfolio_remaining": [r for r in portfolio
                                if _signal(r) in ("Buy", "Fair Value", "Sell", "Strong Sell")],
        "beyond_sb": [r for r in beyond if _signal(r) == "Strong Buy"],
        "beyond_ss": [r for r in beyond if _signal(r) == "Strong Sell"],
        "signal_changes": [r for r in records
                            if (r.get("prev_signal") or "") not in ("", "No Signal", "N/A")
                            and r.get("prev_signal") != _signal(r)],
        "all": records,
    }


def rows_for_full(records):
    """Format records for a 7-column table with signal label."""
    rows = []
    sigs = []
    for r in sorted(records, key=lambda x: -(x.get("value_ratio") or 0)):
        sig = r.get("computed_signal") or ""
        rows.append([
            r.get("company") or r["ticker"],
            r["ticker"],
            r.get("sector") or "",
            fmt_target_pence(r.get("blended_target")),
            fmt_price_pence(r.get("live_price")),
            fmt_vr(r.get("value_ratio")),
            SIGNAL_LABEL.get(sig, sig),
        ])
        sigs.append(sig)
    return rows, sigs


def rows_for_no_signal(records, descending=True):
    """Format records for a 6-column table (no Signal column — group is uniform)."""
    rows = []
    sigs = []
    sorted_recs = sorted(records, key=lambda x: -(x.get("value_ratio") or 0)) if descending \
        else sorted(records, key=lambda x: (x.get("value_ratio") or 99))
    for r in sorted_recs:
        sig = r.get("computed_signal") or ""
        rows.append([
            r.get("company") or r["ticker"],
            r["ticker"],
            r.get("sector") or "",
            fmt_target_pence(r.get("blended_target")),
            fmt_price_pence(r.get("live_price")),
            fmt_vr(r.get("value_ratio")),
        ])
        sigs.append(sig)
    return rows, sigs


def rows_for_signal_changes(records):
    """Format signal-changes records as: Company | Ticker | Prev | New | VR"""
    rows = []
    sigs = []
    for r in records:
        prev = r.get("prev_signal") or ""
        curr = r.get("computed_signal") or ""
        rows.append([
            r.get("company") or r["ticker"],
            r["ticker"],
            r.get("sector") or "",
            SIGNAL_LABEL.get(prev, prev),
            SIGNAL_LABEL.get(curr, curr),
            fmt_vr(r.get("value_ratio")),
        ])
        sigs.append(curr)
    return rows, sigs


def build_heatmap_rows(records):
    """Returns sorted list of (sector, [sb, b, fv, s, ss])."""
    sectors = {}
    for r in records:
        sec = r.get("sector") or "Other"
        sig = r.get("computed_signal") or ""
        if sec not in sectors:
            sectors[sec] = Counter()
        sectors[sec][sig] += 1
    rows = []
    for sec, c in sorted(sectors.items()):
        rows.append((sec, [c.get("Strong Buy", 0), c.get("Buy", 0),
                            c.get("Fair Value", 0), c.get("Sell", 0),
                            c.get("Strong Sell", 0)]))
    return rows


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser()
    parser.add_argument("--date", type=str, default=None,
                        help="Date stamp for output folder (YYYY-MM-DD)")
    parser.add_argument("--out-dir", type=str, default=None,
                        help="Override output directory")
    args = parser.parse_args()

    stamp = args.date or date.today().isoformat()
    out_dir = Path(args.out_dir) if args.out_dir else (DRAFTS / "Article Images" / f"{stamp} tracker_tables")
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output dir: {out_dir}")

    print("Loading records via site loader (matches website universe) …")
    records, held = load_records()
    print(f"  Loaded {len(records)} publishable FTSE equities.")
    print(f"  Held tickers from combined_portfolio.json: {len(held or set())}")

    groups = build_groups(records, held)

    # 01: Portfolio Strong Buy
    rows, sigs = rows_for_full(groups["portfolio_sb"])
    if rows:
        render_table(
            ["Company", "Ticker", "Sector", "Target", "Live Price", "VR", "Signal"],
            rows, sigs, out_dir / "01_portfolio_strong_buy.png",
            col_widths=[0.30, 0.10, 0.18, 0.10, 0.10, 0.07, 0.15],
        )
    else:
        print("  01: no portfolio Strong Buys — skipping")

    # 02: Portfolio Remaining (Buy + FV + Sell + SS)
    rows, sigs = rows_for_full(groups["portfolio_remaining"])
    if rows:
        render_table(
            ["Company", "Ticker", "Sector", "Target", "Live Price", "VR", "Signal"],
            rows, sigs, out_dir / "02_portfolio_remaining.png",
            col_widths=[0.30, 0.10, 0.18, 0.10, 0.10, 0.07, 0.15],
        )
    else:
        print("  02: no portfolio remaining rows — skipping")

    # 03: Beyond portfolio Strong Buys
    rows, sigs = rows_for_no_signal(groups["beyond_sb"])
    if rows:
        render_table(
            ["Company", "Ticker", "Sector", "Target", "Live Price", "VR"],
            rows, sigs, out_dir / "03_beyond_strong_buy.png",
            col_widths=[0.34, 0.12, 0.20, 0.12, 0.12, 0.10],
        )
    else:
        print("  03: no beyond Strong Buys — skipping")

    # 04: Beyond portfolio Strong Sells
    rows, sigs = rows_for_no_signal(groups["beyond_ss"], descending=False)
    if rows:
        render_table(
            ["Company", "Ticker", "Sector", "Target", "Live Price", "VR"],
            rows, sigs, out_dir / "04_beyond_strong_sell.png",
            col_widths=[0.34, 0.12, 0.20, 0.12, 0.12, 0.10],
        )
    else:
        print("  04: no beyond Strong Sells — skipping")

    # 05: Sector heatmap
    heat = build_heatmap_rows(records)
    if heat:
        render_heatmap(heat, out_dir / "05_sector_heatmap.png")
    else:
        print("  05: no heatmap data — skipping")

    # 06: Signal changes (only if any)
    rows, sigs = rows_for_signal_changes(groups["signal_changes"])
    if rows:
        render_table(
            ["Company", "Ticker", "Sector", "Prev", "New", "VR"],
            rows, sigs, out_dir / "06_signal_changes.png",
            col_widths=[0.30, 0.10, 0.18, 0.16, 0.16, 0.10],
        )
        print(f"  06: {len(rows)} signal changes")
    else:
        print("  06: no signal changes this week — image skipped intentionally")

    print()
    print(f"Done. Images saved to: {out_dir}")


if __name__ == "__main__":
    main()
