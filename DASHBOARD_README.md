# Dashboard cheat sheet

## TL;DR

**Normal day:** do nothing. Everything refreshes automatically via scheduled tasks.

**If a dashboard looks stale:** double-click `refresh_all.bat` — one script, runs everything in order.

**Dashboard URL:** https://pcgamer.tail69de8a.ts.net:8787/dashboard2 (Briefing is the home)

---

## Project layout

```
eToro/
├── DASHBOARD_README.md          ← you are here
├── refresh_all.bat              ← double-click to refresh everything on demand
│
├── run_all.py                   ← top-level entry points for the 3 generated dashboards
├── run_combined.py                (also scheduled — you usually don't invoke directly)
├── run_macro.py
├── run_news_server.py           ← the live HTTP server on port 8787 (Tailscale-exposed)
├── run_daily.py  run_on_trade.py  run_tracker.py  (legacy Python runners)
│
├── dashboards/                  ← all HTML dashboard files (generated + live)
│   ├── dashboard2.html              /dashboard2   (Briefing — home)
│   ├── health_dashboard.html        /health
│   ├── bookmarks_dashboard.html     /bookmarks
│   ├── macro_dashboard.html         /macro_dashboard.html
│   ├── eToro_dashboard.html         /eToro_dashboard.html
│   ├── t212_dashboard.html          /t212_dashboard.html
│   ├── Dalkent13_Factsheet.html     /Dalkent13_Factsheet.html
│   └── news_dashboard.html          /  (legacy, unlinked)
│
├── runners/                     ← .bat and .vbs launchers
│   ├── run_news_server.bat / .vbs      (starts the live server)
│   ├── run_calendar_sync.bat / .vbs    (Google Calendar → data/calendar_events.json)
│   ├── run_gmail_sync.bat / .vbs       (Gmail → data/emails.json)
│   ├── run_vault_sync.bat / .vbs       (Obsidian vault → data/health.json)
│   ├── run_earnings_dividends_sync.bat / .vbs
│   ├── run_combined_refresh.bat / .vbs
│   ├── run_macro_refresh.bat / .vbs
│   ├── run_hourly_refresh.bat / .vbs
│   ├── run_dashboard_only.bat          (legacy nightly eToro regen)
│   ├── run_daily.bat / run_daily_briefing.bat
│   └── run_on_trade.bat
│
├── scripts/                     ← Python modules — the real work
│   ├── google_auth.py               shared OAuth helper for Calendar + Gmail
│   ├── sync_calendar.py             hourly Google Calendar sync
│   ├── sync_gmail.py                15-min Gmail sync
│   ├── sync_vault.py                15-min Obsidian vault sync
│   ├── sync_t212.py                 Trading 212 API sync
│   ├── sync_macro.py                Macro series sync (FRED, Yahoo)
│   ├── sync_earnings_dividends.py   Daily Yahoo earnings + dividends
│   ├── generate_dashboard.py        Builds dashboards/eToro_dashboard.html
│   ├── generate_combined_dashboard.py  Builds dashboards/t212_dashboard.html
│   ├── generate_macro_dashboard.py     Builds dashboards/macro_dashboard.html
│   ├── editorial_theme.py           Shared CSS + nav for generated dashboards
│   ├── valuation.py  sync_portfolio.py  reconcile_etoro_closed.py  ...
│   └── ...
│
├── data/                        ← cached JSON + CSV feeding the dashboards
│   ├── calendar_events.json         from sync_calendar.py
│   ├── emails.json                  from sync_gmail.py
│   ├── health.json                  from sync_vault.py
│   ├── earnings.json  dividends.json  from sync_earnings_dividends.py
│   ├── bookmarks.json               ← edit to add/remove bookmark tiles
│   ├── combined_portfolio.json      from generate_combined_dashboard.py
│   ├── etoro_portfolio_output.csv   from the eToro API sync
│   ├── t212_portfolio.json          from sync_t212.py
│   ├── macro.json                   from sync_macro.py
│   └── eToro_Master.xlsx            Source of truth (you edit this)
│
├── credentials/                 ← OAuth tokens (gitignored)
├── logs/                        ← sync + dashboard logs
├── archive/                     ← old/reference stuff
│   (homelab/ moved out — now lives at ../Homelab/ as a separate project)
│
├── etoro.env / t212.env         ← API keys (gitignored)
├── index.html                   ← Substack-style landing page
└── *.svg / *.png                ← logos + images used by Factsheet
```

---

## What refreshes automatically

| Scheduled task            | Cadence       | What it updates                                  |
| ------------------------- | ------------- | ------------------------------------------------ |
| `NewsDashboard`           | on login      | Starts the live news server                      |
| `CalendarSync`            | every hour    | `data/calendar_events.json` → Calendar card      |
| `GmailSync`               | every 15 min  | `data/emails.json` → Email card                  |
| `VaultHealthSync`         | every 15 min  | `data/health.json` ← vault Health notes          |
| `EarningsDividendsSync`   | daily 6am     | `data/earnings.json` + `data/dividends.json`     |
| `eToro Hourly Refresh`    | hourly        | `data/etoro_portfolio_output.csv` + live prices  |
| `eToro_Dashboard_Refresh` | daily 08:00   | `dashboards/eToro_dashboard.html`                |
| `CombinedDashboard_15min` | every 15 min  | `dashboards/t212_dashboard.html` + combined JSON |

The Briefing / Health / Bookmarks pages read from those JSON/CSV files on every load, so data is usually within 15 min of real time.

---

## Pages

| URL                           | Served from                              | Notes                                                 |
| ----------------------------- | ---------------------------------------- | ----------------------------------------------------- |
| `/dashboard2` (Briefing)      | `dashboards/dashboard2.html` (live API)  | Home — news, weather, trains, calendar, sport, email  |
| `/health`                     | `dashboards/health_dashboard.html` (live) | Workouts, PBs, body comp, food log, goals            |
| `/bookmarks`                  | `dashboards/bookmarks_dashboard.html`    | Edit `data/bookmarks.json` to add links               |
| `/macro_dashboard.html`       | `dashboards/macro_dashboard.html`        | Generated — rates, FX, equity indices                 |
| `/eToro_dashboard.html`       | `dashboards/eToro_dashboard.html`        | Generated — eToro portfolio deep dive                 |
| `/t212_dashboard.html`        | `dashboards/t212_dashboard.html`         | Generated — Trading 212 portfolio + activity          |
| `/Dalkent13_Factsheet.html`   | `dashboards/Dalkent13_Factsheet.html`    | Legacy factsheet                                      |

---

## When to run `refresh_all.bat`

- You get a **"Prices stale"** banner on the eToro tab
- You just placed a trade and want immediate numbers
- Something obviously out of date

Running it takes ~30–60s and refreshes everything in order (7 steps).

---

## Troubleshooting

**Server not reachable from phone/tablet**
- Check Tailscale is connected on `pcgamer` (tray icon)
- `tailscale serve status` should list port 8787

**Calendar / Gmail data missing**
- Token expired — run `python scripts\sync_gmail.py` from PowerShell to re-authorize (one browser click)

**Health dashboard empty**
- Vault sync failure — check `logs\vault_sync.log`

**Everything stale**
- Run `refresh_all.bat`

**Scheduled task not found / pointing at legacy path**
- Two tasks (`CombinedDashboard_15min` and `eToro_Dashboard_Refresh`) were registered as SYSTEM so they still reference root paths. Pass-through files at root (`run_combined_refresh_hidden.vbs` + `run_dashboard_only.bat`) forward them to `runners/`. Safe to leave as-is, or run an elevated PowerShell and re-register to point directly at `runners\...`.

---

## Key environment files

- `etoro.env` — Finnhub, NewsAPI, football-data, FRED, football keys
- `t212.env` — Trading 212 API key (+ `.example` stub in repo)
- `credentials/google_calendar_client.json` — OAuth client from Google Cloud Console
- `credentials/google_token.json` — refresh token (created on first `sync_*.py` run)

All gitignored.
