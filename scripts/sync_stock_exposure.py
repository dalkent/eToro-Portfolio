"""sync_stock_exposure.py — Collate stock exposure across eToro, T212, and
the Finances portfolio, including ETF decomposition and pension proxy.

Inputs:
  data/combined_portfolio.json  (eToro + T212 holdings)
  data/finances.json            (Finances ND portfolio — LBG pension, AJ Bell, etc.)

Outputs:
  data/stock_exposure.json
    {
      generated_at, fx_gbpusd,
      total_equity_gbp,
      buckets: { direct, via_funds, via_pension, cash_bonds_crypto },
      funds: [ { ticker, name, total_value_gbp, top_holdings: [ {ticker, name, pct} ] } ],
      consolidated: [ { ticker, name, direct_gbp, indirect_gbp, total_gbp, pct } ],
      unresolved: [ { source, name, value_gbp, reason } ]
    }

Notes:
  - Equity ETFs are decomposed via yfinance `funds_data.top_holdings` (top 10 only,
    so the decomposition is partial — the residual is reported under "rest of fund").
  - LBG Global Equity Fund has no ticker → proxied as a 50/50 blend of IWDA.L +
    VWRL.L top holdings (MSCI World + FTSE All-World).
  - Bond ETFs / bond funds / cash / crypto are tracked in cash_bonds_crypto but
    excluded from the equity exposure roll-up.
"""
from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent.parent
DATA = BASE / "data"
OUT = DATA / "stock_exposure.json"

# ── Fund decomposition config ────────────────────────────────────────────────

# Equity ETFs / mutual funds that we decompose via yfinance top-holdings.
EQUITY_ETFS = {
    "VFEM.L": "Vanguard FTSE Emerging Markets UCITS ETF",
    "INAA.L": "iShares MSCI North America UCITS ETF",
    "VUKE.L": "Vanguard FTSE 100 UCITS ETF",
    "VMID.L": "Vanguard FTSE 250 UCITS ETF",
    "HMXJ.L": "HSBC MSCI Pacific ex Japan UCITS ETF",
    "SAUS.L": "iShares MSCI Australia UCITS ETF",
    "IWDA.L": "iShares Core MSCI World UCITS ETF",  # proxy
    "VWRL.L": "Vanguard FTSE All-World UCITS ETF",  # proxy
}

# Bond / cash ETFs — not decomposed, excluded from equity exposure roll-up.
BOND_ETFS = {
    "LQDE.L", "ISXF.L", "SLXX.L", "VUCP.L", "UKCO.L",
}

# Pension proxy: LBG Global Equity Fund blended as 50% IWDA + 50% VWRL.
PENSION_PROXY_TICKERS = ["IWDA.L", "VWRL.L"]

# Map Finances-portfolio row text fragments → (ticker, name)
# LBG-employee-share rows appear under several labels but all represent LLOY.L.
TEXT_TO_TICKER = [
    (re.compile(r"partnership|deferr[er]d bonus|halifax sharedealing|lloyds banking group", re.I),
     ("LLOY.L", "Lloyds Banking Group plc")),
]

# Ticker normalisation for sheet strings like "NASDAQ:PLTR" or "LON:BP"
_COLON_RE = re.compile(r"^(nasdaq|nyse|lon|lse):", re.I)
# Common name → ticker rescue for rows with no ticker but a known equity name.
NAME_TO_TICKER = {
    "bp plc": "BP.L",
    "shell plc": "SHEL.L",
    "diageo plc": "DGE.L",
    "tate & lyle plc": "TATE.L",
    "j sainsbury plc": "SBRY.L",
    "bitmine immersion technologies inc": "BMNR",
    "novo nordisk a/s": "NVO",
    "coinbase": "COIN",
    "trimble inc": "TRMB",
}


def _norm_ticker(s: str | None) -> str:
    if not s:
        return ""
    s = s.strip()
    s = _COLON_RE.sub("", s)
    # Sheet mixes case like "NASDAq:coi" — upper-case always
    s = s.upper()
    # Treat plain 3-letter LON tickers as .L where appropriate
    if s.endswith(".L"):
        return s
    return s


def _to_l_if_lon(ticker: str, row_ticker_raw: str) -> str:
    """If sheet label started with 'LON:' or 'Lon:', append .L to the ticker."""
    if not ticker:
        return ticker
    if ticker.endswith(".L"):
        return ticker
    if row_ticker_raw and row_ticker_raw.lower().startswith("lon:"):
        # 3-char UK ticker like BP → BP.L
        return ticker + ".L"
    return ticker


def _parse_money_val(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in {"-", "—", "N/A"}:
        return None
    neg = s.startswith("-") or s.startswith("(-")
    s = re.sub(r"[£$€,\s]", "", s.lstrip("-(").rstrip(")"))
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


# ── Fund top-holdings lookup ─────────────────────────────────────────────────

def fetch_fund_holdings(tickers: list[str]) -> dict[str, dict]:
    """Call yfinance for each ETF and return its top-10 holdings."""
    try:
        import yfinance as yf
    except ImportError:
        print("ERROR: yfinance not installed. Run: pip install yfinance")
        return {}
    out = {}
    for t in tickers:
        try:
            tk = yf.Ticker(t)
            df = tk.funds_data.top_holdings
            if df is None or df.empty:
                out[t] = {"name": EQUITY_ETFS.get(t, t), "holdings": [], "error": "empty"}
                continue
            holdings = []
            for symbol, row in df.iterrows():
                holdings.append({
                    "ticker": str(symbol),
                    "name": str(row.get("Name") or symbol),
                    "pct": float(row.get("Holding Percent", 0) or 0),
                })
            out[t] = {
                "name": EQUITY_ETFS.get(t, t),
                "holdings": holdings,
                "covered_pct": sum(h["pct"] for h in holdings),
            }
            time.sleep(0.4)  # be polite to Yahoo
        except Exception as e:  # noqa: BLE001
            out[t] = {"name": EQUITY_ETFS.get(t, t), "holdings": [], "error": str(e)[:120]}
            print(f"  {t}: ERR {e}")
    return out


# ── Source extractors ────────────────────────────────────────────────────────

def _load_combined() -> list[dict]:
    p = DATA / "combined_portfolio.json"
    if not p.exists():
        return []
    d = json.loads(p.read_text(encoding="utf-8"))
    rows = []
    for h in d.get("holdings", []):
        # Prefer the `yahoo` field (already normalised like "BP.L" / "AAPL"),
        # else fall back to the raw ticker.
        ticker = _norm_ticker(h.get("yahoo") or h.get("ticker"))
        # Skip bond ETFs from equity exposure.
        if ticker in BOND_ETFS:
            rows.append({
                "source": h.get("broker", "?"),
                "ticker": ticker,
                "name": h.get("name") or ticker,
                "value_gbp": h.get("value_gbp") or 0,
                "kind": "bond_etf",
            })
            continue
        if ticker == "BTC" or "bitcoin" in (h.get("name") or "").lower():
            rows.append({
                "source": h.get("broker", "?"),
                "ticker": ticker,
                "name": h.get("name") or "Bitcoin",
                "value_gbp": h.get("value_gbp") or 0,
                "kind": "crypto",
            })
            continue
        rows.append({
            "source": h.get("broker", "?"),
            "ticker": ticker,
            "name": h.get("name") or ticker,
            "value_gbp": h.get("value_gbp") or 0,
            "kind": "fund" if ticker in EQUITY_ETFS else "stock",
        })
    return rows


def _load_finances() -> list[dict]:
    p = DATA / "finances.json"
    if not p.exists():
        return []
    d = json.loads(p.read_text(encoding="utf-8"))
    rows_raw = d.get("portfolio", {}).get("rows", [])
    out = []
    # Section tracker (same logic as render_allocation)
    current_section: str | None = None
    stopped = False
    for i, r in enumerate(rows_raw):
        if stopped:
            continue

        def _c(j):
            return (str(r[j]).strip() if j < len(r) and r[j] is not None else "")
        col_a = _c(0); col_b = _c(1); col_c = _c(2); col_e = _c(4); col_g = _c(6)
        a_low = col_a.lower()
        if a_low in ("summary", "overall"):
            stopped = True
            continue
        if a_low.startswith("total "):
            continue
        if col_b.lower() == "investment":
            section_name = col_a
            if not section_name:
                for j in range(i + 1, min(i + 7, len(rows_raw))):
                    cand = (str(rows_raw[j][0]).strip() if len(rows_raw[j]) > 0 and rows_raw[j][0] else "")
                    if cand and cand.lower() not in ("summary", "overall") and not cand.lower().startswith("total "):
                        section_name = cand
                        break
            current_section = section_name or current_section
            continue
        if not col_b:
            continue
        if col_b.lower().startswith("total "):
            continue
        v = _parse_money_val(col_g)
        if v is None or v == 0:
            # Keep rows with value 0 — often placeholder but distorts nothing.
            if v is None:
                continue
        section = current_section or ""

        # LBG Global Equity Fund (pension) — special: no ticker, needs proxy.
        if section.lower() == "pensions":
            name = col_b.lower()
            if "global equity" in name:
                out.append({
                    "source": "LBG Pension",
                    "ticker": "LBG_GLOBAL_EQUITY",
                    "name": col_b,
                    "value_gbp": v,
                    "kind": "pension_equity_fund",
                })
                continue
            if "bond" in name:
                out.append({
                    "source": "LBG Pension",
                    "ticker": "LBG_BOND",
                    "name": col_b,
                    "value_gbp": v,
                    "kind": "bond_fund",
                })
                continue
        # Cash
        if col_c.lower() == "cash" or col_b.lower() == "cash" or "cash" in section.lower():
            out.append({
                "source": section, "ticker": "CASH", "name": col_b or "Cash",
                "value_gbp": v, "kind": "cash",
            })
            continue
        # Crypto
        if col_c.lower() == "crypto" or col_b.lower() == "crypto":
            out.append({
                "source": section, "ticker": "CRYPTO", "name": col_b or "Crypto",
                "value_gbp": v, "kind": "crypto",
            })
            continue
        # Property / Alternative lumps (Seedrs, Bullion, Investing 121, Etoro-alt)
        if "alternative" in section.lower() or col_c.lower() in ("property",):
            out.append({
                "source": section, "ticker": "ALT", "name": col_b,
                "value_gbp": v, "kind": "alternative",
            })
            continue
        # LBG employee share rows (Partnership / Deferred Bonus / Halifax Sharedealing)
        # — all Lloyds Banking Group shares.
        matched_text = None
        for rx, (tk, nm) in TEXT_TO_TICKER:
            if rx.search(col_b):
                matched_text = (tk, nm)
                break
        if matched_text:
            out.append({
                "source": section, "ticker": matched_text[0], "name": matched_text[1],
                "value_gbp": v, "kind": "stock",
            })
            continue

        # Normal ticker from column E
        ticker = _norm_ticker(col_e)
        ticker = _to_l_if_lon(ticker, col_e)
        if not ticker:
            # Try name lookup
            ticker = NAME_TO_TICKER.get(col_b.lower(), "")
        if not ticker:
            out.append({
                "source": section, "ticker": "",
                "name": col_b, "value_gbp": v, "kind": "unresolved",
            })
            continue
        if ticker in BOND_ETFS:
            kind = "bond_etf"
        elif ticker in EQUITY_ETFS:
            kind = "fund"
        else:
            kind = "stock"
        out.append({
            "source": section, "ticker": ticker, "name": col_b,
            "value_gbp": v, "kind": kind,
        })
    return out


# ── Decomposition ────────────────────────────────────────────────────────────

def _decompose_fund(value_gbp: float, fund_data: dict) -> list[dict]:
    """Split a fund position into constituent holdings + residual."""
    out = []
    covered = 0.0
    for h in fund_data.get("holdings", []):
        pct = h["pct"]
        covered += pct
        out.append({
            "ticker": h["ticker"],
            "name": h["name"],
            "value_gbp": value_gbp * pct,
            "via_pct": pct,
        })
    residual_pct = max(1.0 - covered, 0.0)
    if residual_pct > 0.001:
        out.append({
            "ticker": "__REST__",
            "name": "(Rest of fund — not in top 10)",
            "value_gbp": value_gbp * residual_pct,
            "via_pct": residual_pct,
        })
    return out


def _decompose_pension(value_gbp: float, fund_lookup: dict) -> list[dict]:
    """LBG Global Equity Fund proxy = average of IWDA.L + VWRL.L top holdings."""
    per_ticker: dict[str, dict] = {}
    n = 0
    for t in PENSION_PROXY_TICKERS:
        fd = fund_lookup.get(t)
        if not fd or not fd.get("holdings"):
            continue
        n += 1
        for h in fd["holdings"]:
            k = h["ticker"]
            if k not in per_ticker:
                per_ticker[k] = {"ticker": k, "name": h["name"], "pct_sum": 0.0}
            per_ticker[k]["pct_sum"] += h["pct"]
    if n == 0:
        return [{
            "ticker": "__REST__",
            "name": "(LBG Global Equity — proxy unavailable)",
            "value_gbp": value_gbp, "via_pct": 1.0,
        }]
    out = []
    total_pct = 0.0
    for row in per_ticker.values():
        pct = row["pct_sum"] / n
        total_pct += pct
        out.append({
            "ticker": row["ticker"], "name": row["name"],
            "value_gbp": value_gbp * pct, "via_pct": pct,
        })
    residual_pct = max(1.0 - total_pct, 0.0)
    if residual_pct > 0.001:
        out.append({
            "ticker": "__REST__",
            "name": "(Rest of LBG Global Equity Fund — beyond proxy top-10)",
            "value_gbp": value_gbp * residual_pct, "via_pct": residual_pct,
        })
    return out


# ── Consolidation ────────────────────────────────────────────────────────────

_TICKER_ALIASES = {
    # T212 quirks: "DGEl" / "BPl" / "TATEl" / "EZJl" etc. → .L
    # Collapse to same canonical ticker.
}

def _canonical(ticker: str) -> str:
    """Canonicalise a ticker for aggregation. Trust already-normalised tickers
    (e.g., BP.L, AAPL, NVDA) and leave them alone. T212 tickers arrive pre-
    normalised via the `yahoo` field in combined_portfolio.json, so we don't
    need the old 'strip trailing L' heuristic here."""
    if not ticker:
        return ""
    return ticker.upper().strip()


def build_exposure(
    sources: list[dict],
    fund_lookup: dict,
) -> tuple[list[dict], list[dict], list[dict]]:
    """Return (consolidated, funds_summary, unresolved)."""
    # Aggregate: ticker → {direct_gbp, indirect_gbp, total_gbp, name, sources}
    agg: dict[str, dict] = {}

    def _add(ticker: str, name: str, value: float, direct: bool, source: str, via: str = ""):
        canon = _canonical(ticker)
        if canon not in agg:
            agg[canon] = {
                "ticker": canon, "name": name,
                "direct_gbp": 0.0, "indirect_gbp": 0.0,
                "total_gbp": 0.0, "sources": [],
            }
        if direct:
            agg[canon]["direct_gbp"] += value
        else:
            agg[canon]["indirect_gbp"] += value
        agg[canon]["total_gbp"] += value
        # Keep the best human name (prefer longer).
        if len(name) > len(agg[canon]["name"]):
            agg[canon]["name"] = name
        agg[canon]["sources"].append({"source": source, "value_gbp": value, "via": via})

    unresolved: list[dict] = []
    funds_summary: list[dict] = []
    # Cash/bonds/crypto excluded from equity roll-up but returned separately.

    for row in sources:
        kind = row.get("kind")
        value = row.get("value_gbp") or 0.0
        if kind == "stock":
            _add(row["ticker"], row["name"], value, direct=True, source=row["source"])
        elif kind == "fund":
            # Decompose
            fd = fund_lookup.get(row["ticker"])
            if not fd or not fd.get("holdings"):
                unresolved.append({
                    "source": row["source"], "name": row["name"], "ticker": row["ticker"],
                    "value_gbp": value, "reason": (fd or {}).get("error") or "no data",
                })
                continue
            parts = _decompose_fund(value, fd)
            funds_summary.append({
                "ticker": row["ticker"], "name": row["name"] or fd["name"],
                "total_value_gbp": value, "source": row["source"],
                "top_holdings": fd["holdings"],
                "covered_pct": fd.get("covered_pct", 0.0),
            })
            for p in parts:
                if p["ticker"] == "__REST__":
                    _add("__REST_" + row["ticker"], p["name"], p["value_gbp"],
                         direct=False, source=row["source"], via=f"rest of {row['ticker']}")
                else:
                    _add(p["ticker"], p["name"], p["value_gbp"],
                         direct=False, source=row["source"],
                         via=f"via {row['ticker']}")
        elif kind == "pension_equity_fund":
            parts = _decompose_pension(value, fund_lookup)
            funds_summary.append({
                "ticker": "LBG_GLOBAL_EQUITY",
                "name": row["name"] + " (proxy: avg of IWDA.L + VWRL.L)",
                "total_value_gbp": value, "source": row["source"],
                "top_holdings": [{"ticker": p["ticker"], "name": p["name"], "pct": p["via_pct"]}
                                 for p in parts if p["ticker"] != "__REST__"],
                "covered_pct": sum(p["via_pct"] for p in parts if p["ticker"] != "__REST__"),
            })
            for p in parts:
                if p["ticker"] == "__REST__":
                    _add("__REST_LBG", p["name"], p["value_gbp"],
                         direct=False, source=row["source"], via="rest of LBG Global Equity")
                else:
                    _add(p["ticker"], p["name"], p["value_gbp"],
                         direct=False, source=row["source"],
                         via="via LBG Global Equity (proxy)")
        elif kind == "unresolved":
            unresolved.append({
                "source": row["source"], "name": row["name"],
                "ticker": row.get("ticker") or "",
                "value_gbp": value, "reason": "no ticker — needs manual mapping",
            })
        # bond_etf / bond_fund / cash / crypto / alternative — handled outside.

    # Sort consolidated by total value desc
    consolidated = sorted(agg.values(), key=lambda x: x["total_gbp"], reverse=True)
    return consolidated, funds_summary, unresolved


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    print("-- Stock Exposure Sync --")
    combined = _load_combined()
    finances = _load_finances()
    sources = combined + finances
    print(f"  Combined (eToro+T212): {len(combined)} rows")
    print(f"  Finances portfolio:    {len(finances)} rows")

    # Identify all distinct fund tickers we need to fetch.
    fund_tickers_needed = set()
    for row in sources:
        if row.get("kind") == "fund":
            fund_tickers_needed.add(row["ticker"])
        if row.get("kind") == "pension_equity_fund":
            fund_tickers_needed.update(PENSION_PROXY_TICKERS)
    print(f"  Fetching top holdings for {len(fund_tickers_needed)} funds …")
    for t in sorted(fund_tickers_needed):
        print(f"     - {t}")
    fund_lookup = fetch_fund_holdings(sorted(fund_tickers_needed))

    consolidated, funds_summary, unresolved = build_exposure(sources, fund_lookup)

    # Bucket totals
    total_direct = sum(x["direct_gbp"] for x in consolidated)
    total_indirect = sum(x["indirect_gbp"] for x in consolidated)
    total_equity = total_direct + total_indirect
    # Cash / bonds / crypto / alt (for context)
    cash_gbp = sum(r["value_gbp"] for r in sources if r.get("kind") == "cash")
    bond_gbp = sum(r["value_gbp"] for r in sources if r.get("kind") in ("bond_etf", "bond_fund"))
    crypto_gbp = sum(r["value_gbp"] for r in sources if r.get("kind") == "crypto")
    alt_gbp = sum(r["value_gbp"] for r in sources if r.get("kind") == "alternative")

    out = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "totals": {
            "equity_gbp": total_equity,
            "direct_gbp": total_direct,
            "indirect_gbp": total_indirect,
            "cash_gbp": cash_gbp,
            "bonds_gbp": bond_gbp,
            "crypto_gbp": crypto_gbp,
            "alt_gbp": alt_gbp,
            "grand_total_gbp": total_equity + cash_gbp + bond_gbp + crypto_gbp + alt_gbp,
        },
        "consolidated": [
            {**x, "pct_of_equity": (x["total_gbp"] / total_equity * 100) if total_equity else 0}
            for x in consolidated
        ],
        "funds": funds_summary,
        "unresolved": unresolved,
        "pension_proxy": {
            "description": "LBG Global Equity Fund approximated as 50/50 IWDA.L + VWRL.L top holdings",
            "tickers": PENSION_PROXY_TICKERS,
        },
    }
    OUT.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\n  Wrote {OUT}")
    print(f"  Total equity exposure: £{total_equity:,.0f} "
          f"(direct £{total_direct:,.0f} + indirect £{total_indirect:,.0f})")
    print(f"  Unresolved positions:  {len(unresolved)}")
    print(f"  Distinct tickers:      {len(consolidated)}")


if __name__ == "__main__":
    main()
