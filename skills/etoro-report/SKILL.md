# eToro Buy/Sell Report Generator

## When to Use This Skill
Trigger whenever the user says something like:
- "Lloyds Sell"
- "IMB Buy"
- "LLOY.L Sell"
- "Rolls Royce Buy"
- "[any company name or ticker] [Buy or Sell]"

---

## Data File

**Canonical source: JSON** — `C:\Users\Neil\My Drive\eToro Sync\etoro_master.json`

Use this first. It's atomic, parses cleanly, and is the same source the website
and tracker read from. The Excel `eToro_Master.xlsx` is the upstream input that
`valuation.py` writes into; downstream readers (this skill included) should not
read the xlsx directly unless the JSON is unavailable.

**Fallback (rare):** `C:\Users\Neil\ClaudeCode\eToro\data\eToro_Master.xlsx`

---

## Step-by-Step Process

### Step 1 — Parse the User's Input
Extract two things:
- **Company / Ticker**: The stock they named (e.g., "Lloyds", "LLOY.L", "Rolls Royce")
- **Action**: Buy or Sell

### Step 2 — Read the data

Preferred: parse the JSON.

```python
import json
with open(r"C:\Users\Neil\My Drive\eToro Sync\etoro_master.json", encoding="utf-8") as f:
    raw = f.read()
try:
    data = json.loads(raw)
except json.JSONDecodeError:
    # Tolerate trailing junk from older non-atomic writes
    data = json.JSONDecoder().raw_decode(raw)[0]

# Per-ticker valuations (replaces the Assumptions sheet read below):
valuations = data["assumptions"]["valuations"]
# Portfolio rows (replaces the Portfolio sheet read below):
portfolio = data["sheets"]["portfolio"]["objects"]
# Watchlist rows (replaces the Watchlist sheet read below):
watchlist = data["sheets"]["watchlist"]["objects"]
```

Fallback only when the JSON cache is missing or corrupt: use Python with `openpyxl`
(`data_only=True`) to read `eToro_Master.xlsx` per the column maps below.

**Assumptions tab** (header row = Row 7, data starts Row 8):
```
Col 0:  Ticker
Col 1:  Company
Col 2:  Sector
Col 3:  Beta
Col 4:  WACC
Col 5:  g1 (5yr Growth)
Col 6:  g2 (Terminal)
Col 7:  Target Price (GBP/USD)  ← may be None if not set
Col 8:  Val 1 (DCF / Banks:DDM)
Col 9:  Val 2 (DDM / Banks:P/B / AM:P/B)
Col 10: Val 3 (EPV / Fin:Earn Cap / GI:P/TB)
Col 11: Blended Target (GBP / USD)
Col 12: Model / Method
Col 13: Last Updated
Col 14: Notes
Col 15: Prev Signal
Col 16: Current Signal
```

**Portfolio tab** (header row = Row 2, data starts Row 3):
```
Col 1:  Company Name
Col 2:  Sector
Col 3:  eToro Ticker
Col 5:  Currency  (GBp = pence, USD = dollars)
Col 8:  Units Held
Col 9:  Avg Buy Price (Local)
Col 10: Invested (USD)
Col 11: Live Price (Local)
Col 16: Capital ROI %
Col 17: Div 2023 (USD)
Col 18: Div 2024 (USD)
Col 19: Div 2025 (USD)
Col 20: Div 2026E (USD)
Col 21: Total Divs (USD)
Col 22: Div Return %
Col 24: ROI (with Divs)
Col 27: Target (GBP/USD)
Col 28: Live Price (GBP/USD)
Col 29: Value Ratio
Col 30: Signal
```

**Watchlist tab** (header row = Row 2, data starts Row 3):
```
Col 1: Company / Name
Col 2: Sector
Col 3: eToro Ticker
Col 12: Signal
```

**Tickers tab** (header row = Row 2, data starts Row 3):
```
Col 1: Company / Name
Col 2: FTSE Ticker
Col 7: Sector
Col 8: Asset Type
Col 9: In Portfolio (Yes/No)
Col 10: In Watchlist (Yes/No)
```

### Step 3 — Find Peers
Use the **Watchlist tab** (col 2 = Sector) to find 2–3 stocks in the same sector.
Do NOT use the Tickers tab for sector matching — it uses broader categories.
For each peer, pull their Current Signal and blended target from the Assumptions tab.
Exclude the subject stock itself. Aim for the 2–3 most prominent names.

### Step 4 — Fetch Live Price via Web Search
Do NOT rely on the Excel for the current price — it only reflects the last saved value.
Instead, search: `"[TICKER] share price today pence"` or `"[TICKER] stock price"`
Extract the current price in the stock's native currency (pence for GBp stocks, USD for US stocks).
Use this live price for all calculations below.

### Step 5 — Calculate ROI (if stock is in Portfolio)
Using data from the Portfolio tab and the live price from Step 4:

```python
# For GBp stocks:
gbp_usd = 1.27  # from Assumptions tab row 3, col 1 (update if changed)

current_value_usd = units * (live_price_p / 100) * gbp_usd
capital_pl        = current_value_usd - invested_usd
capital_roi_pct   = capital_pl / invested_usd * 100

div_total = sum(d for d in [div_2023, div_2024, div_2025, div_2026e] if d)
total_roi_pct = (capital_pl + div_total) / invested_usd * 100

# Premium/discount to target:
target_p = blended_target_gbp * 100  # convert GBP to pence
premium_pct = (live_price_p - target_p) / target_p * 100
# positive = trading above target (overvalued), negative = below (undervalued)
```

### Step 6 — Search for Recent News
Search: `"[Company Name] news 2026"` and `"[Company Name] results [year]"`
Find 2–3 recent, relevant items — results, analyst moves, regulatory news, macro catalysts.

### Step 7 — Generate the Report
Follow the template below exactly. Use the pre-converted Unicode bold strings for headers.

---

## Report Template

Use this template, substituting real data from the Excel file and web search.

```
🔴 SELL | [COMPANY NAME] | $[ETORO_TICKER] | [SECTOR]    ← use 🟢 BUY for buys
(note: only the coloured circle emoji at the start — no other emojis in the report)

𝗪𝗵𝗮𝘁 𝗶𝘀 [Company]?
[2-3 sentences: what the company does, its main markets/products, its size or prominence
(FTSE 100 / FTSE 250 etc.), and one distinguishing characteristic. Write naturally —
no em dashes, no colons mid-sentence where a comma would do.]

𝗠𝘆 𝗩𝗲𝗿𝗱𝗶𝗰𝘁
[1-2 sentences in first person. Reference your target price vs current price and the
signal. E.g. "My target price for [Company] is Xp, against a current price of around
Xp. At that premium to fair value and with a Strong Sell signal, I am closing this
position." No em dashes — use commas or full stops to break up the sentence instead.]

𝗩𝗮𝗹𝘂𝗮𝘁𝗶𝗼𝗻
Valuation method: [Model/Method from Assumptions col 12 — written naturally, e.g.
"DDM, Price/Book and EPV" for banks, "DCF, DDM and EPV" for standard stocks]
My target price: [BlendedTarget converted to pence or USD, 2dp]
Signal: [Signal text only — no emoji, just plain text e.g. "Strong Sell"]

𝗠𝘆 𝗣𝗼𝘀𝗶𝘁𝗶𝗼𝗻   ← INCLUDE ONLY IF THE STOCK IS IN THE PORTFOLIO TAB
Capital ROI: [Capital ROI %]% | Total ROI including dividends: [ROI with Divs %]%
(DO NOT include invested amount in dollars or dividend amounts in dollars — ROI percentages only)

𝗪𝗵𝘆 𝗜'𝗺 𝗦𝗲𝗹𝗹𝗶𝗻𝗴   ← or 𝗪𝗵𝘆 𝗜'𝗺 𝗕𝘂𝘆𝗶𝗻𝗴 depending on action
• [Valuation reason: how far current price is above/below your target price, expressed
  as a % premium or discount. No em dashes. Use "at a X% premium to my target" etc.]
• [Fundamental or sector reason: macro, competitive, regulatory, structural]
• [News or catalyst reason: reference a specific recent news item from Step 4]
• [Optional 4th reason if genuinely useful: dividend trend, high beta risk, etc.]

𝗥𝗲𝗰𝗲𝗻𝘁 𝗡𝗲𝘄𝘀
• [News item 1, brief and factual, source and approximate date in brackets]
• [News item 2, brief and factual]
• [News item 3, optional]

𝗣𝗲𝗲𝗿𝘀 𝘁𝗼 𝗪𝗮𝘁𝗰𝗵
$[PEER1_TICKER] | [Peer Company Name] | Signal: [signal] | [one-line reason why relevant]
$[PEER2_TICKER] | [Peer Company Name] | Signal: [signal] | [one-line reason]
[Optional 3rd peer on same format]
(Use pipe character | as separator — not em dashes)

Not financial advice. These are my personal views based on my own valuation models.
Always do your own research before investing.
#[TickerHashtag] #FTSE100 #[SectorHashtag] #ValueInvesting
```

---

## Reference: eToro Bold Headers (Pre-Converted)
Copy these exactly into the report — do not re-process them:

| Section | Bold Unicode |
|---------|-------------|
| What is [X]? | 𝗪𝗵𝗮𝘁 𝗶𝘀 |
| My Verdict | 𝗠𝘆 𝗩𝗲𝗿𝗱𝗶𝗰𝘁 |
| Valuation | 𝗩𝗮𝗹𝘂𝗮𝘁𝗶𝗼𝗻 |
| My Position | 𝗠𝘆 𝗣𝗼𝘀𝗶𝘁𝗶𝗼𝗻 |
| Why I'm Selling | 𝗪𝗵𝘆 𝗜'𝗺 𝗦𝗲𝗹𝗹𝗶𝗻𝗴 |
| Why I'm Buying | 𝗪𝗵𝘆 𝗜'𝗺 𝗕𝘂𝘆𝗶𝗻𝗴 |
| Recent News | 𝗥𝗲𝗰𝗲𝗻𝘁 𝗡𝗲𝘄𝘀 |
| Peers to Watch | 𝗣𝗲𝗲𝗿𝘀 𝘁𝗼 𝗪𝗮𝘁𝗰𝗵 |

---

## Reference: Signal Text (plain text only — no emojis)
| Signal from Excel | Display in report |
|--------|---------|
| Strong Buy | Strong Buy |
| Buy | Buy |
| Fair Value | Fair Value |
| Sell | Sell |
| Strong Sell | Strong Sell |

## Reference: Action Indicator (opening line only)
| Action | Opening |
|--------|-------------|
| SELL | 🔴 SELL |
| BUY | 🟢 BUY |

---

## Formatting Rules
- NO em dashes (never use —). Use commas, full stops, or pipe | instead.
- NO emojis except the single 🔴 or 🟢 on the opening line.
- NO dollar amounts for invested capital or dividends received. ROI percentages only.
- Show ONLY the blended target price (col 11), not individual Val 1/2/3 prices.
  Call it "My target price" — not "blended target" or "fair value".
- Mention the valuation method (col 12) in plain English but do not show per-method prices.
- Currency: if Currency col = "GBp" convert blended target to pence (multiply by 100, show as Xp).
  If "USD" show as $X.XX.
- Round all prices to 2 decimal places.
- If the stock is NOT in the Portfolio tab, omit "My Position" entirely.
- Capital ROI % and ROI with Divs % may be None if the live price hasn't been fetched yet.
  If None, omit the My Position section rather than showing null values.
- Keep the whole report under 2,800 characters for eToro compatibility.
- Hashtags at the end: ticker, FTSE100, sector in CamelCase (e.g. #FinancialServices), #ValueInvesting.

---

## Python Snippet for Reading Data (xlsx fallback only)

Prefer the JSON snippet in Step 2 above. The xlsx code below is for when the
JSON cache is unavailable.

```python
import openpyxl

wb = openpyxl.load_workbook(
    r"C:\Users\Neil\ClaudeCode\eToro\data\eToro_Master.xlsx",
    data_only=True
)

# Assumptions: rows[6] = header, rows[7:] = data
assumptions_rows = list(wb['Assumptions'].iter_rows(values_only=True))
assumptions_header = assumptions_rows[6]
assumptions_data = assumptions_rows[7:]

# Portfolio: rows[1] = header, rows[2:] = data
portfolio_rows = list(wb['Portfolio'].iter_rows(values_only=True))
portfolio_header = portfolio_rows[1]
portfolio_data = portfolio_rows[2:]

# Watchlist: rows[1] = header, rows[2:] = data
watchlist_rows = list(wb['Watchlist'].iter_rows(values_only=True))
watchlist_data = watchlist_rows[2:]

# Search by ticker or partial company name (case-insensitive)
query = "lloy"  # from user input
match = next(
    (r for r in assumptions_data
     if r[0] and query.lower() in str(r[0]).lower()
     or r[1] and query.lower() in str(r[1]).lower()),
    None
)
```
