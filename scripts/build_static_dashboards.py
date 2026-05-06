"""build_static_dashboards.py — produce standalone copies of every dashboard
suitable for hosting on Google Drive (no Docker / Tailscale / news-server
required to view them).

The generated dashboards in `dashboards/` come in two flavours:

  • Already-static (eToro, T212, macro, finances, exposure, factsheet) — they
    just need copying to the Drive Upload folder.

  • Live-fetch (dashboard2 = Briefing, health, bookmarks) — these call
    /api/news, /api/health, /api/bookmarks at runtime. To make them work
    offline, this script bakes the latest data into a small fetch-interceptor
    block that overrides window.fetch for the matched URLs.

Output:
  dashboards/<name>.html -> unchanged (in-place; the live-app keeps working too)
  C:\\Users\\Neil\\My Drive\\Upload\\<name>.html -> standalone, openable from
                                                 anywhere, no server needed.

Usage:
  python scripts/build_static_dashboards.py            # full build
  python scripts/build_static_dashboards.py --dry-run  # show what would happen

The fetch interceptor is small JS that:
  - Looks up the URL (path + query) in window.__STATIC_API_DATA
  - Returns a Promise<Response> with the embedded JSON
  - Falls through to the network only if no static match was found

This keeps the dashboard JS unchanged — same fetch('/api/health') call in
the page, just intercepted before it hits the network.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR     = Path(__file__).parent.parent
DASHBOARDS   = BASE_DIR / "dashboards"
UPLOAD_DIR   = Path(r"C:\Users\Neil\My Drive\Upload")
DATA_DIR     = BASE_DIR / "data"

# Make run_news_server importable (it sits at the repo root)
sys.path.insert(0, str(BASE_DIR))


# ── Live-fetch dashboards: declare the API -> data-producer mapping ────────────
def _safe_call(label: str, fn, default):
    """Call a data-producer function with a clear error if it fails."""
    try:
        return fn()
    except Exception as e:  # noqa: BLE001
        print(f"  WARN: {label} failed — {type(e).__name__}: {e}")
        return default


def build_static_payloads() -> dict[str, dict]:
    """Returns a map of dashboard_filename -> { url_pattern: data_object }.

    The url_pattern is matched against the request's URL path (+ query string).
    The data_object is whatever the corresponding /api/* endpoint would return.
    """
    print("  Importing run_news_server …")
    from run_news_server import (
        build_news,
        health_data,
        football_standings,
        FOOTBALL_COMPETITIONS,
    )

    payloads: dict[str, dict] = {}

    # ── health_dashboard.html  ->  /api/health ─────────────────────────────────
    print("  Building /api/health …")
    health = _safe_call("health_data", health_data, {})
    payloads["health_dashboard.html"] = {"/api/health": health}

    # ── bookmarks_dashboard.html  ->  /api/bookmarks ───────────────────────────
    print("  Building /api/bookmarks …")
    bm_path = DATA_DIR / "bookmarks.json"
    bookmarks = json.loads(bm_path.read_text(encoding="utf-8")) if bm_path.exists() else {}
    payloads["bookmarks_dashboard.html"] = {"/api/bookmarks": bookmarks}

    # ── dashboard2.html (Briefing)  ->  /api/news + /api/football per comp ─────
    print("  Building /api/news (this hits live RSS / Finnhub — may take ~30s) …")
    news = _safe_call("build_news", lambda: build_news(force=False), {})
    briefing = {"/api/news": news}
    print("  Building /api/football for each competition …")
    for code, label in FOOTBALL_COMPETITIONS:
        try:
            data = football_standings(code)
            # Store under the exact URL the page generates: /api/football?competition=<code>
            briefing[f"/api/football?competition={code}"] = data
            print(f"     {code} ({label}): {len(data.get('table') or data.get('standings') or [])} rows")
        except Exception as e:  # noqa: BLE001
            print(f"     {code}: failed — {e}")
            briefing[f"/api/football?competition={code}"] = {"error": str(e)}
    payloads["dashboard2.html"] = briefing

    return payloads


# ── HTML splicing ─────────────────────────────────────────────────────────────
INTERCEPTOR_JS_TEMPLATE = """
<script id="__static_api_interceptor">
/* Static-mode fetch interceptor — built {generated_at} by build_static_dashboards.py.
   Overrides window.fetch so /api/* calls return baked-in data instead of hitting
   the news-server. Other URLs (CDN fonts, etc.) pass through untouched. */
window.__STATIC_API_DATA = {payload_json};
(function(){{
  if (!window.__STATIC_API_DATA) return;
  var DATA = window.__STATIC_API_DATA;
  var origFetch = window.fetch;
  window.fetch = function(input, init) {{
    var url = (typeof input === 'string') ? input : (input && input.url) || '';
    /* Try direct match (path + query) */
    if (url && Object.prototype.hasOwnProperty.call(DATA, url)) {{
      return Promise.resolve(_resp(DATA[url]));
    }}
    /* Try parsed path + query */
    try {{
      var u = new URL(url, window.location.origin);
      var pq = u.pathname + (u.search || '');
      if (Object.prototype.hasOwnProperty.call(DATA, pq)) {{
        return Promise.resolve(_resp(DATA[pq]));
      }}
      if (Object.prototype.hasOwnProperty.call(DATA, u.pathname)) {{
        return Promise.resolve(_resp(DATA[u.pathname]));
      }}
    }} catch (e) {{}}
    return origFetch.apply(this, arguments);
  }};
  function _resp(obj){{
    return new Response(JSON.stringify(obj), {{
      status: 200, headers: {{'Content-Type': 'application/json'}}
    }});
  }}
}})();
</script>
"""


# Map of absolute server-side URLs → relative .html filenames for the
# Drive Upload copies. The live news server routes things like /dashboard2 to
# dashboard2.html; Google Drive doesn't, so when viewing the static copies
# from Drive (especially on tablet), nav links must be relative .html paths.
NAV_HREF_REWRITES = {
    "/dashboard2":                   "dashboard2.html",
    "/health":                       "health_dashboard.html",
    "/bookmarks":                    "bookmarks_dashboard.html",
    "/exposure":                     "exposure_dashboard.html",
    "/macro_dashboard.html":         "macro_dashboard.html",
    "/eToro_dashboard.html":         "eToro_dashboard.html",
    "/t212_dashboard.html":          "t212_dashboard.html",
    "/finances_dashboard.html":      "finances_dashboard.html",
    "/exposure_dashboard.html":      "exposure_dashboard.html",
    "/combined_dashboard.html":      "combined_dashboard.html",
    "/Dalkent13_Factsheet.html":     "Dalkent13_Factsheet.html",
    "/Dalkent13_Factsheet_mobile.html": "Dalkent13_Factsheet_mobile.html",
}


def rewrite_nav_links(html: str) -> str:
    """Rewrite absolute server-side nav hrefs to relative .html filenames.

    Touches only href="..." attributes — leaves other absolute paths alone.
    Idempotent (safe to run twice).
    """
    for old, new in NAV_HREF_REWRITES.items():
        html = html.replace(f'href="{old}"', f'href="{new}"')
    return html


def staticify_html(html: str, payload: dict, generated_at: str) -> str:
    """Inject the fetch-interceptor block right after <head>."""
    # Build the script block
    script_block = INTERCEPTOR_JS_TEMPLATE.format(
        generated_at=generated_at,
        payload_json=json.dumps(payload, ensure_ascii=False),
    )
    # Inject after the opening <head> tag (handles attributes too)
    head_end = html.find("<head")
    if head_end == -1:
        return script_block + "\n" + html
    head_end = html.find(">", head_end)
    if head_end == -1:
        return script_block + "\n" + html
    return html[: head_end + 1] + "\n" + script_block + html[head_end + 1 :]


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="Build standalone dashboards for Drive")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print plan without writing/copying anything")
    parser.add_argument("--no-network", action="store_true",
                        help="Skip live-fetch dashboards (skip building /api/news payloads)")
    args = parser.parse_args()

    print("=" * 56)
    print(f"  build_static_dashboards.py — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 56)

    if not DASHBOARDS.exists():
        sys.exit(f"ERROR: {DASHBOARDS} not found")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

    # Build payloads for live-fetch dashboards (calls live RSS / football APIs)
    payloads: dict[str, dict] = {}
    if not args.no_network:
        try:
            payloads = build_static_payloads()
        except Exception as e:  # noqa: BLE001
            print(f"  WARN: failed to build live payloads: {e}")
            print(f"  Live-fetch dashboards will NOT be staticified this run.")
            payloads = {}

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    copied = 0
    staticified = 0
    skipped = 0

    for src in sorted(DASHBOARDS.glob("*.html")):
        name = src.name
        dest = UPLOAD_DIR / name
        if name in payloads:
            # Live-fetch dashboard — staticify it (bake data) AND rewrite nav links
            html = src.read_text(encoding="utf-8")
            html = staticify_html(html, payloads[name], generated_at)
            html = rewrite_nav_links(html)
            if args.dry_run:
                print(f"  [dry] STATICIFY {name} -> {dest}  ({len(html):,} bytes)")
            else:
                dest.write_text(html, encoding="utf-8")
                staticified += 1
                print(f"  STATICIFY {name} -> {dest}  ({len(html):,} bytes)")
        else:
            # Already-static dashboard — read, rewrite nav links, write to Drive Upload.
            # (Can't shutil.copy2 anymore since we mutate the contents.)
            if args.dry_run:
                print(f"  [dry] COPY      {name} -> {dest}")
            else:
                html = src.read_text(encoding="utf-8")
                html = rewrite_nav_links(html)
                dest.write_text(html, encoding="utf-8")
                copied += 1
                print(f"  COPY      {name} -> {dest}")

    # Clean up stale Upload entries that no longer have a source
    print()
    print("  Stale Upload entries (no source dashboard):")
    valid = {p.name for p in DASHBOARDS.glob("*.html")}
    # Keep these legacy items even if they're stale — user may rely on them
    KEEP_STALE = {"index.html", "eToro_Dashboard_Refresh.xml"}
    for f in UPLOAD_DIR.glob("*.html"):
        if f.name not in valid and f.name not in KEEP_STALE:
            print(f"    !  {f.name} (last modified {datetime.fromtimestamp(f.stat().st_mtime):%Y-%m-%d %H:%M})")

    print()
    print(f"  Done: {copied} copied · {staticified} staticified · {skipped} skipped")
    if not args.dry_run:
        print(f"  All dashboards available at {UPLOAD_DIR}")


if __name__ == "__main__":
    main()
