"""News dashboard server.

Serves a multi-section news dashboard (Local, International, Markets,
Portfolio, Sport) with a refresh button. Reads tickers from
data/etoro_portfolio_output.csv.

Run:  python run_news_server.py
Open: http://127.0.0.1:8787
"""
import csv
import json
import os
import re
import time
import webbrowser
import http.cookiejar
import threading
from concurrent.futures import ThreadPoolExecutor
from urllib.request import HTTPCookieProcessor, build_opener
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo
from html import escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
ENV_FILE = os.path.join(ROOT, "etoro.env")
DASHBOARDS_DIR = os.path.join(ROOT, "dashboards")
HTML_FILE = os.path.join(DASHBOARDS_DIR, "dashboard2.html")

if os.path.exists(ENV_FILE):
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())

NEWSAPI_KEY = os.environ.get("NEWSAPI_KEY", "")
FINNHUB_KEY = os.environ.get("FINNHUB_API_KEY", "")
FOOTBALL_DATA_KEY = os.environ.get("FOOTBALL_DATA_KEY", "")
PORT = int(os.environ.get("NEWS_PORT", "8787"))
CACHE_TTL = 600

_cache: dict[str, tuple[float, object]] = {}


def fetch(url: str, timeout: int = 15) -> bytes:
    req = Request(url, headers={"User-Agent": "Mozilla/5.0 (news-dashboard)"})
    with urlopen(req, timeout=timeout) as r:
        return r.read()


def strip_html(s: str) -> str:
    return re.sub(r"<[^>]+>", "", s or "").replace("&nbsp;", " ").strip()


def parse_rss(xml_bytes: bytes, limit: int = 15, default_source: str = "") -> list[dict]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return []
    items = []
    for it in root.iter("item"):
        src_el = it.find("source")
        src = ""
        if src_el is not None and src_el.text:
            src = src_el.text.strip()
        items.append({
            "title": (it.findtext("title") or "").strip(),
            "link": (it.findtext("link") or "").strip(),
            "snippet": strip_html(it.findtext("description") or "")[:280],
            "published": (it.findtext("pubDate") or "").strip(),
            "source": src or default_source,
        })
        if len(items) >= limit:
            break
    return items


def newsapi_headlines(**params) -> list[dict]:
    if not NEWSAPI_KEY:
        return []
    params["apiKey"] = NEWSAPI_KEY
    qs = "&".join(f"{k}={quote(str(v))}" for k, v in params.items())
    url = f"https://newsapi.org/v2/top-headlines?{qs}"
    try:
        data = json.loads(fetch(url))
    except Exception:
        return []
    if data.get("status") != "ok":
        return []
    out = []
    for a in data.get("articles", [])[:15]:
        out.append({
            "title": a.get("title") or "",
            "link": a.get("url") or "",
            "snippet": (a.get("description") or "")[:280],
            "published": a.get("publishedAt") or "",
            "source": (a.get("source") or {}).get("name") or "",
        })
    return out


def finnhub_company_news(symbol: str, limit: int = 4) -> list[dict]:
    if not FINNHUB_KEY:
        return []
    today = datetime.now(timezone.utc).date()
    frm = today.replace(day=1).isoformat()
    to = today.isoformat()
    url = (f"https://finnhub.io/api/v1/company-news"
           f"?symbol={quote(symbol)}&from={frm}&to={to}&token={FINNHUB_KEY}")
    try:
        data = json.loads(fetch(url))
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    data.sort(key=lambda a: a.get("datetime", 0), reverse=True)
    out = []
    for a in data[:limit]:
        ts = a.get("datetime")
        pub = ""
        if isinstance(ts, (int, float)) and ts > 0:
            pub = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%d %b %Y")
        out.append({
            "title": a.get("headline") or "",
            "link": a.get("url") or "",
            "snippet": (a.get("summary") or "")[:280],
            "published": pub,
            "source": a.get("source") or "",
        })
    return out


def google_news_rss(query: str, limit: int = 4) -> list[dict]:
    url = (f"https://news.google.com/rss/search?q={quote(query)}"
           f"&hl=en-GB&gl=GB&ceid=GB:en")
    try:
        return parse_rss(fetch(url), limit=limit)
    except Exception:
        return []


CRYPTO_NAMES = {
    "BTC": "Bitcoin",
    "ETH": "Ethereum",
    "SOL": "Solana",
    "XRP": "XRP",
    "ADA": "Cardano",
    "DOGE": "Dogecoin",
}


def load_tickers(n: int = 12) -> list[dict]:
    path = os.path.join(DATA, "etoro_portfolio_output.csv")
    if not os.path.exists(path):
        return []
    rows = []
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            t = (row.get("Ticker") or "").strip()
            name = (row.get("Company_Name") or "").strip()
            sector = (row.get("Sector") or "").strip()
            try:
                val = float(row.get("Current_Value_USD") or 0)
            except ValueError:
                val = 0
            if not t or t.upper() == "CASH" or val <= 0:
                continue
            if sector.lower() == "crypto" and t.upper() in CRYPTO_NAMES:
                name = CRYPTO_NAMES[t.upper()]
            try:
                units = float(row.get("Units_Held") or 0)
            except ValueError:
                units = 0
            rows.append({"ticker": t, "name": name, "sector": sector, "value": val, "units": units})
    rows.sort(key=lambda r: r["value"], reverse=True)
    return rows[:n]


def load_portfolio_and_watchlist_tickers() -> list[dict]:
    """Union of portfolio + watchlist tickers for broader sweeps (dividends,
    earnings). Portfolio tickers keep their held units/value; watchlist entries
    have units=0 and value=0 (so est_payment drops to None).

    Reads watchlist from data/etoro_master.json (produced by sync_xlsx_to_vault).
    """
    out = list(load_tickers(n=500))  # portfolio, full list
    seen = {r["ticker"].upper() for r in out}

    master_path = os.path.join(DATA, "etoro_master.json")
    if not os.path.exists(master_path):
        return out
    try:
        with open(master_path, encoding="utf-8") as f:
            master = json.load(f)
    except Exception:
        return out
    for obj in (master.get("sheets") or {}).get("watchlist", {}).get("objects") or []:
        ticker = str(obj.get("eToro Ticker") or obj.get("Yahoo Ticker") or "").strip()
        if not ticker or ticker.upper() in seen:
            continue
        seen.add(ticker.upper())
        name = str(obj.get("Company / Name") or ticker).strip()
        sector = str(obj.get("Sector") or "").strip()
        out.append({
            "ticker": ticker,
            "name":   name,
            "sector": sector,
            "value":  0.0,
            "units":  0.0,
            "watchlist": True,
        })
    return out


def finnhub_symbol(ticker: str) -> str:
    # LSE tickers: LLOY.L -> LON:LLOY (Finnhub uses exchange prefix for intl)
    if ticker.upper().endswith(".L"):
        return "LON:" + ticker[:-2].upper()
    return ticker.upper()


def ticker_news(ticker: str, name: str, sector: str = "") -> dict:
    if sector.lower() == "crypto":
        items = google_news_rss(f"{name} cryptocurrency")
    else:
        items = finnhub_company_news(finnhub_symbol(ticker))
        if not items:
            q = f'"{name}"' if name else ticker
            items = google_news_rss(q + " stock")
    return {"ticker": ticker, "name": name, "items": items}


def cached(key: str, producer) -> object:
    now = time.time()
    hit = _cache.get(key)
    if hit and now - hit[0] < CACHE_TTL:
        return hit[1]
    data = producer()
    # Don't cache empty results — retry on next request instead of serving stale empty.
    if data or (isinstance(data, dict) and data):
        _cache[key] = (now, data)
    return data


def local_news() -> list[dict]:
    items = []
    for url in [
        "https://feeds.bbci.co.uk/news/uk/rss.xml",
        "https://feeds.bbci.co.uk/news/politics/rss.xml",
    ]:
        try:
            items.extend(parse_rss(fetch(url), limit=10, default_source="BBC News"))
        except Exception:
            pass
    seen, out = set(), []
    for it in items:
        if it["link"] in seen:
            continue
        seen.add(it["link"])
        out.append(it)
    return out[:15]


def international_news() -> list[dict]:
    feeds = [
        ("https://feeds.bbci.co.uk/news/world/rss.xml", "BBC News"),
        ("https://www.theguardian.com/world/rss", "The Guardian"),
        ("https://www.aljazeera.com/xml/rss/all.xml", "Al Jazeera"),
        ("https://rss.nytimes.com/services/xml/rss/nyt/World.xml", "New York Times"),
    ]
    items = []
    for url, source in feeds:
        try:
            items.extend(parse_rss(fetch(url), limit=8, default_source=source))
        except Exception:
            pass
    seen, out = set(), []
    for it in items:
        if it["link"] in seen:
            continue
        seen.add(it["link"])
        out.append(it)
    return out[:20]


def markets_news() -> list[dict]:
    items = newsapi_headlines(category="business", country="gb")
    if not items:
        items = newsapi_headlines(category="business", language="en")
    if items:
        return items
    try:
        return parse_rss(fetch("https://feeds.bbci.co.uk/news/business/rss.xml"), limit=15)
    except Exception:
        return []


def sport_news() -> list[dict]:
    try:
        return parse_rss(fetch("https://feeds.bbci.co.uk/sport/rss.xml"), limit=15, default_source="BBC Sport")
    except Exception:
        return []


MAX_AGE_DAYS = 3


def filter_recent(items: list[dict], days: int = MAX_AGE_DAYS) -> list[dict]:
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    return [it for it in items if parse_story_time(it.get("published", "")) >= cutoff]


_STOPWORDS = {"the", "and", "for", "with", "from", "that", "this", "what",
              "have", "has", "says", "after", "over", "live", "updates",
              "video", "watch", "photos", "analysis", "explainer", "news"}


def _title_signature(title: str) -> str:
    t = re.sub(r"[-|–:]\s*(live|reuters|bbc news|guardian|al jazeera|nyt|ap).*$", "", title, flags=re.I)
    t = re.sub(r"[^a-z0-9\s]", " ", t.lower())
    words = [w for w in t.split() if len(w) >= 4 and w not in _STOPWORDS]
    return " ".join(words[:5])


def dedupe_stories(items: list[dict]) -> list[dict]:
    seen, out = set(), []
    for it in items:
        sig = _title_signature(it.get("title", ""))
        if not sig or sig in seen:
            if sig:
                continue
        seen.add(sig)
        out.append(it)
    return out


WMO_CODES = {
    0: ("☀", "Clear"), 1: ("🌤", "Mainly clear"),
    2: ("⛅", "Partly cloudy"), 3: ("☁", "Overcast"),
    45: ("🌫", "Foggy"), 48: ("🌫", "Freezing fog"),
    51: ("🌦", "Light drizzle"), 53: ("🌦", "Drizzle"), 55: ("🌧", "Heavy drizzle"),
    61: ("🌧", "Light rain"), 63: ("🌧", "Rain"), 65: ("🌧", "Heavy rain"),
    71: ("🌨", "Light snow"), 73: ("🌨", "Snow"), 75: ("❄", "Heavy snow"),
    77: ("🌨", "Snow grains"),
    80: ("🌦", "Rain showers"), 81: ("🌧", "Heavy showers"), 82: ("⛈", "Violent showers"),
    85: ("🌨", "Snow showers"), 86: ("❄", "Heavy snow showers"),
    95: ("⛈", "Thunderstorm"), 96: ("⛈", "Thunder + hail"), 99: ("⛈", "Severe storm"),
}


def weather() -> dict:
    url = ("https://api.open-meteo.com/v1/forecast"
           "?latitude=51.374&longitude=0.097"
           "&current=temperature_2m,apparent_temperature,weather_code,wind_speed_10m,"
           "relative_humidity_2m,precipitation"
           "&daily=temperature_2m_max,temperature_2m_min,weather_code,sunrise,sunset,"
           "precipitation_probability_max,uv_index_max"
           "&timezone=Europe%2FLondon&forecast_days=4")
    try:
        data = json.loads(fetch(url))
    except Exception:
        return {}
    cur = data.get("current") or {}
    daily = data.get("daily") or {}
    cur_icon, cur_text = WMO_CODES.get(cur.get("weather_code"), ("", ""))
    days = []
    times = daily.get("time") or []
    for i, day in enumerate(times[:4]):
        code = (daily.get("weather_code") or [None] * 4)[i]
        icon, text = WMO_CODES.get(code, ("", ""))
        days.append({
            "day": datetime.fromisoformat(day).strftime("%a"),
            "icon": icon, "text": text,
            "max": round((daily.get("temperature_2m_max") or [0] * 4)[i]),
            "min": round((daily.get("temperature_2m_min") or [0] * 4)[i]),
            "precip": (daily.get("precipitation_probability_max") or [0] * 4)[i] or 0,
        })
    sunrise = (daily.get("sunrise") or [""])[0]
    sunset = (daily.get("sunset") or [""])[0]
    def _t(iso):
        try:
            return datetime.fromisoformat(iso).strftime("%H:%M")
        except (ValueError, TypeError):
            return ""
    uv = (daily.get("uv_index_max") or [None])[0]
    return {
        "location": "Orpington",
        "current": {
            "temp": round(cur.get("temperature_2m") or 0),
            "feels_like": round(cur.get("apparent_temperature") or 0),
            "icon": cur_icon, "text": cur_text,
            "wind": round(cur.get("wind_speed_10m") or 0),
            "humidity": round(cur.get("relative_humidity_2m") or 0),
            "precip_now": cur.get("precipitation") or 0,
        },
        "sunrise": _t(sunrise),
        "sunset": _t(sunset),
        "uv_max": round(uv) if isinstance(uv, (int, float)) else None,
        "forecast": days,
    }


CRYPTO_TICKERS = {"BTC", "ETH", "SOL", "XRP", "ADA", "DOGE", "MATIC", "DOT"}
BOND_TICKERS = {"LQDE.L", "BND", "TLT", "IBTM.L", "IGLT.L", "GILS.L", "AGG"}


def _classify_holding(h: dict) -> str:
    t = (h.get("ticker") or "").upper()
    base = t.replace(".L", "").replace("_EQ", "").rstrip("LD")
    if t in CRYPTO_TICKERS or base in CRYPTO_TICKERS:
        return "Crypto"
    if t in BOND_TICKERS or "BOND" in (h.get("name") or "").upper():
        return "Bonds"
    return "Equities"


def portfolio_summary() -> dict:
    path = os.path.join(DATA, "combined_portfolio.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            combined = json.load(f)
    except Exception:
        return {}
    holdings = combined.get("holdings") or []
    if not holdings:
        return {}
    total_value = combined.get("total_value_gbp") or sum(h.get("value_gbp") or 0 for h in holdings)
    total_pnl = combined.get("total_pnl_gbp") or sum(h.get("pnl_gbp") or 0 for h in holdings)
    assets_value = sum(h.get("value_gbp") or 0 for h in holdings) or 1
    top = sorted(holdings, key=lambda h: h.get("value_gbp") or 0, reverse=True)[:7]
    by_broker = combined.get("by_broker") or {}

    buckets: dict[str, float] = {"Equities": 0.0, "Crypto": 0.0, "Bonds": 0.0, "Cash": 0.0}
    for h in holdings:
        cls = _classify_holding(h)
        buckets[cls] = buckets.get(cls, 0) + (h.get("value_gbp") or 0)
    buckets["Cash"] = sum((b.get("cash_gbp") or 0) for b in by_broker.values())
    total_with_cash = sum(buckets.values()) or 1
    allocation = [
        {"name": name, "value": round(val, 2), "pct": val / total_with_cash * 100}
        for name, val in buckets.items() if val > 0
    ]
    allocation.sort(key=lambda x: x["pct"], reverse=True)

    return {
        "currency": "GBP",
        "total_value": round(total_value, 2),
        "total_pnl": round(total_pnl, 2),
        "positions": combined.get("positions") or len(holdings),
        "by_broker": {
            k: {
                "value": round(v.get("value_gbp") or 0, 2),
                "pnl":   round(v.get("pnl_gbp") or 0, 2),
                "positions": v.get("positions") or 0,
                "cash":  round(v.get("cash_gbp") or 0, 2),
            }
            for k, v in by_broker.items()
        },
        "allocation": allocation,
        "top_holdings": [
            {
                "ticker": h.get("ticker") or "",
                "name":   h.get("name") or "",
                "broker": h.get("broker") or "",
                "value":  round(h.get("value_gbp") or 0, 2),
                "weight": round((h.get("value_gbp") or 0) / assets_value * 100, 1),
                "pnl":    round(h.get("pnl_gbp") or 0, 2),
                "roi":    round(h.get("roi") or 0, 2),
            }
            for h in top
        ],
        "history": [],
        "generated_at": combined.get("generated_at") or "",
    }


def economic_calendar() -> list[dict]:
    if not FINNHUB_KEY:
        return []
    today = datetime.now(timezone.utc).date()
    end = today + timedelta(days=7)
    url = (f"https://finnhub.io/api/v1/calendar/economic"
           f"?from={today}&to={end}&token={FINNHUB_KEY}")
    try:
        data = json.loads(fetch(url))
    except Exception:
        return []
    events = (data.get("economicCalendar") or []) if isinstance(data, dict) else []
    filtered = [e for e in events if e.get("country") in ("GB", "US")]
    filtered.sort(key=lambda e: (e.get("time") or ""))
    out = []
    for e in filtered[:10]:
        out.append({
            "event": e.get("event", ""),
            "country": e.get("country", ""),
            "time": e.get("time", ""),
            "impact": e.get("impact", ""),
            "estimate": e.get("estimate"),
            "previous": e.get("prev"),
        })
    return out


def trains_from(crs: str, dest: str, dest_label: str) -> list[dict]:
    url = f"https://huxley2.azurewebsites.net/departures/{crs}/to/{dest}/5"
    try:
        data = json.loads(fetch(url))
    except Exception:
        return []
    services = (data.get("trainServices") or []) if isinstance(data, dict) else []
    out = []
    for s in services[:4]:
        out.append({
            "std": s.get("std", ""),
            "etd": s.get("etd", ""),
            "dest": dest_label,
            "platform": s.get("platform") or "-",
            "operator": s.get("operator", ""),
        })
    return out


def train_departures() -> dict:
    legs = [
        # direction, label, from_crs, to_crs, destination_name
        ("from", "London Bridge",     "ORP", "LBG", "London Bridge"),
        ("from", "Cannon St",         "ORP", "CST", "Cannon St"),
        ("from", "Charing Cross",     "ORP", "CHX", "Charing Cross"),
        ("to",   "from London Bridge","LBG", "ORP", "Orpington"),
        ("to",   "from Cannon St",    "CST", "ORP", "Orpington"),
        ("to",   "from Charing Cross","CHX", "ORP", "Orpington"),
    ]
    def _fetch(leg):
        direction, label, src, dst, dest_label = leg
        return {
            "direction": direction,
            "label":     label,
            "services":  trains_from(src, dst, dest_label),
        }
    with ThreadPoolExecutor(max_workers=6) as ex:
        results = list(ex.map(_fetch, legs))
    return {
        "from": [r for r in results if r["direction"] == "from"],
        "to":   [r for r in results if r["direction"] == "to"],
    }


INDEX_SYMBOLS = [
    ("FTSE 100", "^FTSE"),
    ("S&P 500", "^GSPC"),
    ("Nasdaq", "^IXIC"),
    ("GBP/USD", "GBPUSD=X"),
    ("Bitcoin", "BTC-USD"),
    ("Gold", "GC=F"),
    ("Brent", "BZ=F"),
    ("WTI", "CL=F"),
    ("Nat Gas", "NG=F"),
    ("Copper", "HG=F"),
]


def yahoo_quote(symbol: str) -> dict:
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{quote(symbol)}?interval=1d&range=2d")
    try:
        data = json.loads(fetch(url))
    except Exception:
        return {}
    result = (((data.get("chart") or {}).get("result") or [{}])[0]) or {}
    meta = result.get("meta") or {}
    price = meta.get("regularMarketPrice")
    prev = meta.get("chartPreviousClose") or meta.get("previousClose")
    if price is None or prev is None or prev == 0:
        return {}
    pct = (price - prev) / prev * 100
    return {"price": price, "change": price - prev, "pct": pct}


def market_indices() -> list[dict]:
    out = []
    with ThreadPoolExecutor(max_workers=6) as ex:
        quotes = list(ex.map(lambda s: (s[0], s[1], yahoo_quote(s[1])), INDEX_SYMBOLS))
    for label, symbol, q in quotes:
        if not q:
            continue
        out.append({
            "label": label,
            "symbol": symbol,
            "price": round(q["price"], 2 if q["price"] < 1000 else 0),
            "pct": round(q["pct"], 2),
        })
    return out


FOOTBALL_COMPETITIONS = [
    ("PL", "Premier League"),
    ("ELC", "Championship"),
    ("PD", "La Liga"),
    ("BL1", "Bundesliga"),
    ("SA", "Serie A"),
    ("FL1", "Ligue 1"),
    ("CL", "Champions Lg"),
]


def football_standings(code: str = "PL") -> dict:
    if not FOOTBALL_DATA_KEY:
        return {}
    url = f"https://api.football-data.org/v4/competitions/{code}/standings"
    try:
        req = Request(url, headers={
            "X-Auth-Token": FOOTBALL_DATA_KEY,
            "User-Agent": "Mozilla/5.0 (news-dashboard)",
        })
        with urlopen(req, timeout=15) as r:
            data = json.loads(r.read())
    except Exception:
        return {}
    standings = data.get("standings") or []
    table = next((s for s in standings if s.get("type") == "TOTAL"), None)
    if not table:
        return {}
    rows = []
    for e in (table.get("table") or [])[:20]:
        team = e.get("team") or {}
        rows.append({
            "pos": e.get("position"),
            "team": team.get("shortName") or team.get("name") or "",
            "crest": team.get("crest") or "",
            "played": e.get("playedGames"),
            "won": e.get("won"),
            "drawn": e.get("draw"),
            "lost": e.get("lost"),
            "gd": e.get("goalDifference"),
            "pts": e.get("points"),
            "form": e.get("form") or "",
        })
    return {
        "code": code,
        "name": (data.get("competition") or {}).get("name", ""),
        "matchday": (data.get("season") or {}).get("currentMatchday"),
        "rows": rows,
    }


def _fetch_fixtures_for_comp(code: str, date: str) -> list[dict]:
    url = (f"https://api.football-data.org/v4/competitions/{code}/matches"
           f"?dateFrom={date}&dateTo={date}")
    try:
        req = Request(url, headers={
            "X-Auth-Token": FOOTBALL_DATA_KEY,
            "User-Agent": "Mozilla/5.0 (news-dashboard)",
        })
        with urlopen(req, timeout=15) as r:
            data = json.loads(r.read())
    except Exception:
        return []
    return (data.get("matches") or []) if isinstance(data, dict) else []


FIXTURE_KEEP_COMPS = {"PL", "CL"}
FIXTURE_KEEP_TEAMS = {"manchester united", "man united", "england"}


def _fixture_is_kept(comp: str, home: str, away: str) -> bool:
    if comp in FIXTURE_KEEP_COMPS:
        return True
    h, a = (home or "").lower(), (away or "").lower()
    for t in FIXTURE_KEEP_TEAMS:
        if t in h or t in a:
            return True
    return False


def earnings_from_json() -> list[dict]:
    path = os.path.join(DATA, "earnings.json")
    if not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8") as f:
            return (json.load(f) or {}).get("items") or []
    except Exception:
        return []


def dividends_from_json() -> list[dict]:
    path = os.path.join(DATA, "dividends.json")
    if not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8") as f:
            return (json.load(f) or {}).get("items") or []
    except Exception:
        return []


def health_data() -> dict:
    path = os.path.join(DATA, "health.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def emails_data() -> dict:
    path = os.path.join(DATA, "emails.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}
    msgs = data.get("messages") or []
    # Relative time helper done in JS; just pass through.
    out = []
    for m in msgs[:10]:
        out.append({
            "from":     m.get("from") or "",
            "subject":  m.get("subject") or "",
            "snippet":  (m.get("snippet") or "")[:140],
            "received": m.get("received") or "",
            "unread":   bool(m.get("unread")),
            "link":     m.get("link") or "https://mail.google.com/",
        })
    return {
        "generated_at": data.get("generated_at") or "",
        "unread_total": int(data.get("unread_total") or 0),
        "messages":     out,
    }


def calendar_events() -> dict:
    path = os.path.join(DATA, "calendar_events.json")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}
    now = datetime.now(ZoneInfo("Europe/London"))
    today = now.date()
    tomorrow = today + timedelta(days=1)
    groups: dict[str, list] = {"today": [], "tomorrow": []}
    for e in data.get("events", []):
        start_str = e.get("start", "")
        if not start_str:
            continue
        try:
            if "T" in start_str:
                s = datetime.fromisoformat(start_str).astimezone(ZoneInfo("Europe/London"))
                d = s.date()
                time_str = s.strftime("%H:%M")
            else:
                d = datetime.strptime(start_str, "%Y-%m-%d").date()
                time_str = "All day"
        except ValueError:
            continue
        bucket = "today" if d == today else ("tomorrow" if d == tomorrow else None)
        if not bucket:
            continue
        groups[bucket].append({
            "summary": e.get("summary", ""),
            "time":    time_str,
            "all_day": e.get("all_day", False),
            "link":    e.get("link", ""),
            "location": e.get("location") or "",
        })
    for b in groups.values():
        b.sort(key=lambda x: (x.get("all_day", False), x.get("time", "")))
    return {
        "generated_at": data.get("generated_at", ""),
        "today":    groups["today"],
        "tomorrow": groups["tomorrow"],
    }


def fixtures_today() -> list[dict]:
    if not FOOTBALL_DATA_KEY:
        return []
    today = datetime.now(ZoneInfo("Europe/London")).date().isoformat()
    # Fetch broad set so we catch Man United + England in any competition the free tier covers.
    codes = ["PL", "CL", "ELC", "PD", "BL1", "SA", "FL1", "EC", "WC"]
    matches: list[dict] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        for ms in ex.map(lambda c: _fetch_fixtures_for_comp(c, today), codes):
            matches.extend(ms)
    # Apply keep-filter and dedupe by (home, away, utcDate)
    seen = set()
    kept = []
    for m in matches:
        comp_code = (m.get("competition") or {}).get("code", "")
        home_name = ((m.get("homeTeam") or {}).get("shortName")
                     or (m.get("homeTeam") or {}).get("name") or "")
        away_name = ((m.get("awayTeam") or {}).get("shortName")
                     or (m.get("awayTeam") or {}).get("name") or "")
        if not _fixture_is_kept(comp_code, home_name, away_name):
            continue
        key = (home_name, away_name, m.get("utcDate", ""))
        if key in seen:
            continue
        seen.add(key)
        kept.append(m)
    matches = kept
    out = []
    for m in matches:
        comp = (m.get("competition") or {}).get("code", "")
        home = (m.get("homeTeam") or {}).get("shortName") or (m.get("homeTeam") or {}).get("tla") or ""
        away = (m.get("awayTeam") or {}).get("shortName") or (m.get("awayTeam") or {}).get("tla") or ""
        status = m.get("status", "")
        score = (m.get("score") or {}).get("fullTime") or {}
        utc_date = m.get("utcDate") or ""
        local_time = ""
        sort_key = utc_date
        if utc_date:
            try:
                dt = datetime.fromisoformat(utc_date.replace("Z", "+00:00"))
                local_time = dt.astimezone(ZoneInfo("Europe/London")).strftime("%H:%M")
            except ValueError:
                pass
        out.append({
            "competition": comp,
            "home": home, "away": away,
            "status": status,
            "home_score": score.get("home"),
            "away_score": score.get("away"),
            "time": local_time,
            "_sort": sort_key,
        })
    out.sort(key=lambda x: x.get("_sort") or "")
    for o in out:
        o.pop("_sort", None)
    return out


_yahoo_jar: http.cookiejar.CookieJar | None = None
_yahoo_crumb: str = ""
_yahoo_lock = threading.Lock()


def _yahoo_auth() -> bool:
    global _yahoo_jar, _yahoo_crumb
    with _yahoo_lock:
        if _yahoo_crumb:
            return True
        jar = http.cookiejar.CookieJar()
        opener = build_opener(HTTPCookieProcessor(jar))
        opener.addheaders = [("User-Agent", "Mozilla/5.0")]
        try:
            opener.open("https://fc.yahoo.com/", timeout=10)
        except Exception:
            pass
        try:
            r = opener.open("https://query1.finance.yahoo.com/v1/test/getcrumb", timeout=10)
            _yahoo_crumb = r.read().decode()
            _yahoo_jar = jar
            return bool(_yahoo_crumb)
        except Exception:
            return False


def yahoo_earnings_date(symbol: str) -> dict:
    if not _yahoo_auth() or _yahoo_jar is None:
        return {}
    url = (f"https://query1.finance.yahoo.com/v10/finance/quoteSummary/{quote(symbol)}"
           f"?modules=calendarEvents&crumb={quote(_yahoo_crumb)}")
    opener = build_opener(HTTPCookieProcessor(_yahoo_jar))
    opener.addheaders = [("User-Agent", "Mozilla/5.0")]
    try:
        r = opener.open(url, timeout=10)
        data = json.loads(r.read())
    except Exception:
        return {}
    try:
        earn = data["quoteSummary"]["result"][0]["calendarEvents"]["earnings"]
    except (KeyError, IndexError, TypeError):
        return {}
    dates = earn.get("earningsDate") or []
    if not dates:
        return {}
    first = dates[0] if isinstance(dates[0], dict) else {}
    fmt = first.get("fmt")
    ts = first.get("raw")
    if not fmt and ts:
        fmt = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
    if not fmt:
        return {}
    eps_avg = (earn.get("earningsAverage") or {}).get("raw")
    return {"date": fmt, "eps_est": eps_avg}


def yahoo_dividend_info(symbol: str) -> dict:
    if not _yahoo_auth() or _yahoo_jar is None:
        return {}
    url = (f"https://query1.finance.yahoo.com/v10/finance/quoteSummary/{quote(symbol)}"
           f"?modules=summaryDetail&crumb={quote(_yahoo_crumb)}")
    opener = build_opener(HTTPCookieProcessor(_yahoo_jar))
    opener.addheaders = [("User-Agent", "Mozilla/5.0")]
    try:
        r = opener.open(url, timeout=10)
        data = json.loads(r.read())
    except Exception:
        return {}
    try:
        sd = data["quoteSummary"]["result"][0]["summaryDetail"]
    except (KeyError, IndexError, TypeError):
        return {}
    ex_div = sd.get("exDividendDate") or {}
    rate = sd.get("dividendRate") or {}
    yld = sd.get("dividendYield") or {}
    ex_fmt = ex_div.get("fmt") if isinstance(ex_div, dict) else None
    rate_raw = rate.get("raw") if isinstance(rate, dict) else None
    yld_raw = yld.get("raw") if isinstance(yld, dict) else None
    if not ex_fmt or not rate_raw:
        return {}
    return {"ex_date": ex_fmt, "rate": rate_raw, "yield": yld_raw}


def dividend_calendar() -> list[dict]:
    port = load_portfolio_and_watchlist_tickers()
    if not port:
        return []
    with ThreadPoolExecutor(max_workers=10) as ex:
        results = list(ex.map(
            lambda t: (t, yahoo_dividend_info(t["ticker"])),
            port,
        ))
    today = datetime.now(timezone.utc).date()
    oldest = today - timedelta(days=60)
    out = []
    for t, info in results:
        if not info:
            continue
        try:
            d = datetime.strptime(info["ex_date"], "%Y-%m-%d").date()
        except ValueError:
            continue
        if d < oldest:
            continue
        units = t.get("units", 0) or 0
        est_payment = info["rate"] * units if units else None
        out.append({
            "ticker":      t["ticker"],
            "name":        t["name"],
            "ex_date":     info["ex_date"],
            "rate":        info["rate"],
            "yield":       info.get("yield"),
            "est_payment": est_payment,
            "upcoming":    d >= today,
            "watchlist":   bool(t.get("watchlist")),
        })
    out.sort(key=lambda x: x["ex_date"], reverse=True)
    return out[:40]


def earnings_calendar() -> list[dict]:
    port = load_portfolio_and_watchlist_tickers()
    if not port:
        return []
    today = datetime.now(timezone.utc).date()
    cutoff = today + timedelta(days=60)
    with ThreadPoolExecutor(max_workers=6) as ex:
        results = list(ex.map(
            lambda t: (t, yahoo_earnings_date(t["ticker"])),
            port,
        ))
    out = []
    for t, er in results:
        if not er or not er.get("date"):
            continue
        try:
            d = datetime.strptime(er["date"], "%Y-%m-%d").date()
        except ValueError:
            continue
        if d < today or d > cutoff:
            continue
        out.append({
            "ticker": t["ticker"],
            "name": t["name"],
            "date": er["date"],
            "hour": "",
            "eps_est": er.get("eps_est"),
            "quarter": None,
        })
    out.sort(key=lambda e: e["date"])
    return out[:15]


def trending(geo: str = "GB") -> list[dict]:
    try:
        xml = fetch(f"https://trends.google.com/trending/rss?geo={geo}")
    except Exception:
        return []
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    ns = {"ht": "https://trends.google.com/trending/rss"}
    out = []
    for it in root.iter("item"):
        traffic_el = it.find("ht:approx_traffic", ns)
        related = []
        for ni in it.findall("ht:news_item", ns):
            t = ni.find("ht:news_item_title", ns)
            u = ni.find("ht:news_item_url", ns)
            s = ni.find("ht:news_item_source", ns)
            related.append({
                "title": strip_html(t.text if t is not None and t.text else ""),
                "link": (u.text or "").strip() if u is not None and u.text else "",
                "source": (s.text or "").strip() if s is not None and s.text else "",
            })
            if len(related) >= 2:
                break
        out.append({
            "term": (it.findtext("title") or "").strip(),
            "traffic": (traffic_el.text or "").strip() if traffic_el is not None and traffic_el.text else "",
            "related": related,
        })
        if len(out) >= 15:
            break
    return out


def portfolio_news() -> list[dict]:
    tickers = load_tickers()
    with ThreadPoolExecutor(max_workers=6) as ex:
        return list(ex.map(
            lambda t: ticker_news(t["ticker"], t["name"], t.get("sector", "")),
            tickers,
        ))


def parse_story_time(s: str) -> datetime:
    if not s:
        return datetime.min.replace(tzinfo=timezone.utc)
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        try:
            dt = parsedate_to_datetime(s)
        except (ValueError, TypeError):
            return datetime.min.replace(tzinfo=timezone.utc)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def build_latest(sections: dict, limit: int = 6) -> list[dict]:
    pool = []
    for key in ("local", "international", "markets", "sport"):
        for it in sections.get(key) or []:
            pool.append({**it, "category": key})
    pool.sort(key=lambda x: parse_story_time(x.get("published", "")), reverse=True)
    seen, out = set(), []
    for it in pool:
        if it["link"] in seen:
            continue
        seen.add(it["link"])
        out.append(it)
        if len(out) >= limit:
            break
    return out


def build_news(force: bool = False) -> dict:
    if force:
        _cache.clear()
    sections = {
        "local": dedupe_stories(filter_recent(cached("local", local_news))),
        "international": dedupe_stories(filter_recent(cached("international", international_news))),
        "markets": dedupe_stories(filter_recent(cached("markets", markets_news))),
        "sport": dedupe_stories(filter_recent(cached("sport", sport_news))),
        "portfolio": [
            {**g, "items": filter_recent(g.get("items") or [])}
            for g in cached("portfolio", portfolio_news)
        ],
        "trending_uk": cached("trending_uk", lambda: trending("GB")),
        "trending_us": cached("trending_us", lambda: trending("US")),
        "weather": cached("weather", weather),
        "portfolio_summary": portfolio_summary(),
        "economic_calendar": cached("economic_calendar", economic_calendar),
        "trains": train_departures(),
        "market_indices": cached("market_indices", market_indices),
        "football": cached("football_PL", lambda: football_standings("PL")),
        "football_competitions": [{"code": c, "name": n} for c, n in FOOTBALL_COMPETITIONS],
        "fixtures": cached("fixtures_today", fixtures_today),
        "earnings":  earnings_from_json(),
        "dividends": dividends_from_json(),
        "calendar": calendar_events(),
        "emails": emails_data(),
    }
    sections["latest"] = build_latest(sections)
    sections["generated"] = time.strftime("%d %b %Y %H:%M")
    return sections


ALLOWED_STATIC = {
    "macro_dashboard.html", "combined_dashboard.html", "t212_dashboard.html",
    "eToro_dashboard.html", "Dalkent13_Factsheet.html",
    "Dalkent13_Factsheet_mobile.html",
    "dashboard2.html", "health_dashboard.html", "bookmarks_dashboard.html",
    "finances_dashboard.html", "exposure_dashboard.html",
}
STATIC_MIMES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".svg": "image/svg+xml", ".gif": "image/gif", ".webp": "image/webp",
    ".ico": "image/x-icon", ".json": "application/json; charset=utf-8",
}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a, **kw):
        pass

    def _send(self, status: int, body: bytes, content_type: str, no_store: bool = False):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        if no_store:
            self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _serve_static(self, filename: str) -> bool:
        # Look in dashboards/ first, then root (for images / legacy assets).
        path = os.path.join(DASHBOARDS_DIR, filename)
        if not os.path.isfile(path):
            path = os.path.join(ROOT, filename)
        if not os.path.isfile(path):
            return False
        ext = os.path.splitext(filename)[1].lower()
        mime = STATIC_MIMES.get(ext, "application/octet-stream")
        with open(path, "rb") as f:
            body = f.read()
        self._send(200, body, mime)
        return True

    def do_GET(self):
        p = urlparse(self.path)
        if p.path in ("/", "/index.html"):
            try:
                with open(HTML_FILE, "rb") as f:
                    body = f.read()
            except FileNotFoundError:
                self.send_error(500, "dashboard2.html missing")
                return
            self._send(200, body, "text/html; charset=utf-8")
            return
        if p.path == "/dashboard2":
            if self._serve_static("dashboard2.html"):
                return
        if p.path == "/health":
            if self._serve_static("health_dashboard.html"):
                return
        if p.path == "/api/health":
            body = json.dumps(health_data()).encode("utf-8")
            self._send(200, body, "application/json; charset=utf-8", no_store=True)
            return
        if p.path == "/bookmarks":
            if self._serve_static("bookmarks_dashboard.html"):
                return
        if p.path == "/exposure":
            if self._serve_static("exposure_dashboard.html"):
                return
        if p.path == "/api/bookmarks":
            bm_path = os.path.join(DATA, "bookmarks.json")
            try:
                with open(bm_path, "rb") as f:
                    body = f.read()
                self._send(200, body, "application/json; charset=utf-8", no_store=True)
            except FileNotFoundError:
                self._send(200, b"{}", "application/json; charset=utf-8", no_store=True)
            return
        # Static dashboards + their image/svg assets
        clean = p.path.lstrip("/")
        if clean in ALLOWED_STATIC:
            if self._serve_static(clean):
                return
        if any(clean.lower().endswith(ext) for ext in (".png", ".svg", ".jpg", ".jpeg", ".ico", ".webp", ".gif")):
            if "/" not in clean and self._serve_static(clean):
                return
        if p.path == "/api/news":
            force = "refresh" in p.query
            data = build_news(force=force)
            self._send(200, json.dumps(data).encode("utf-8"),
                       "application/json; charset=utf-8", no_store=True)
        elif p.path == "/api/football":
            from urllib.parse import parse_qs
            qs = parse_qs(p.query)
            code = (qs.get("competition") or ["PL"])[0].upper()
            valid = {c for c, _ in FOOTBALL_COMPETITIONS}
            if code not in valid:
                code = "PL"
            data = cached(f"football_{code}", lambda: football_standings(code))
            self._send(200, json.dumps(data).encode("utf-8"),
                       "application/json; charset=utf-8", no_store=True)
        else:
            self.send_error(404)


def main():
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://127.0.0.1:{PORT}"
    print(f"News dashboard running at {url} — Ctrl+C to stop")
    if not os.environ.get("NEWS_NO_BROWSER"):
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")


if __name__ == "__main__":
    main()
