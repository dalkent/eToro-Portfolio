"""Shared editorial theme CSS for dashboard generators.

Applies the warm editorial palette (cream bg, orange-red accent, green up),
Fraunces serif headlines, Inter sans, JetBrains Mono numbers.

Classes covered across generators:
  - Base: body, h1/h2, p
  - Layout: .card, .section, .grid, .grid-cards, .two-col, .three-col
  - Portfolio/KPI: .kpi-grid, .kpi, .kpi-value, .kpi-label, .kpi-sub, .card.grand
  - Chips/badges: .chip, .badge, .region, .stale, .stale-warning, .warn
  - Tables: th, td, .num, .delta
  - Movement: .pos, .neg
  - eToro-specific: .header, .header-right, .nav-bar, .chart-wrap, .signal-grid
  - Combined-specific: .broker-bar
"""

FONTS_LINK = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">'
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
    '<link href="https://fonts.googleapis.com/css2?'
    'family=Fraunces:opsz,wght@9..144,300;9..144,400;9..144,500;9..144,600&'
    'family=Inter:wght@300;400;500;600;700&'
    'family=JetBrains+Mono:wght@300;400;500;600&display=swap" rel="stylesheet">'
)

_NAV_ITEMS = [
    ("briefing", "/dashboard2", "Briefing"),
    ("health", "/health", "Health"),
    ("bookmarks", "/bookmarks", "Bookmarks"),
    ("macro", "/macro_dashboard.html", "Macro"),
    ("etoro", "/eToro_dashboard.html", "eToro"),
    ("t212", "/t212_dashboard.html", "T212"),
    ("finances", "/finances_dashboard.html", "Finances"),
    ("exposure", "/exposure_dashboard.html", "Exposure"),
    ("factsheet", "/Dalkent13_Factsheet.html", "Factsheet"),
]


def nav_html(active: str = "", brand: str = "", privacy: bool = False) -> str:
    """Return shared editorial nav bar HTML. Pass active=slug to highlight a link.

    When privacy=True, also renders a "Private" toggle that blurs £/$/€ values.
    Useful on money-bearing dashboards (eToro, T212, Finances).
    """
    links = "".join(
        f'<a href="{href}"' + (' class="active"' if active == slug else '') + f'>{label}</a>'
        for slug, href, label in _NAV_ITEMS
    )
    brand_html = f'<h1 class="ed-nav-brand">{brand}</h1>' if brand else ''
    privacy_btn = (
        '<button class="ed-privacy-toggle" id="ed-privacy-toggle" '
        'aria-label="Toggle privacy mode (hide money values)">'
        '<span id="ed-privacy-label">Private</span></button>'
    ) if privacy else ''
    return (
        f'<div class="ed-nav">'
        f'{brand_html}'
        f'<div class="ed-nav-links">{links}</div>'
        f'<div class="ed-nav-controls">'
        f'{privacy_btn}'
        f'<button class="ed-theme-toggle" id="ed-theme-toggle" aria-label="Toggle theme">'
        f'<span id="ed-theme-label">Dark</span></button>'
        f'</div></div>'
    )


THEME_JS = r"""
<script>
(function(){
  /* ── Theme toggle ─────────────────────────────────────────── */
  var KEY = 'ed-theme';
  function apply(mood){
    document.documentElement.dataset.mood = mood;
    var l = document.getElementById('ed-theme-label');
    if (l) l.textContent = mood === 'dark' ? 'Light' : 'Dark';
  }
  var saved = localStorage.getItem(KEY) || 'warm';
  apply(saved);
  document.addEventListener('click', function(e){
    if (e.target.closest('#ed-theme-toggle')){
      var next = (document.documentElement.dataset.mood === 'dark') ? 'warm' : 'dark';
      localStorage.setItem(KEY, next);
      apply(next);
    }
  });

  /* ── Privacy mode: blur all £/$/€ numeric values ─────────── */
  var PKEY = 'ed-privacy';
  /* Matches £1,234  £1,234.56  -£12.50  £(1.2)  $1,200  €500k  £1.5m  etc. */
  var MONEY_RE = /-?\(?\s?[£$€]\s?-?\d[\d,]*(?:\.\d+)?\s?[kKmMbB]?\)?/g;
  /* Bare number inside a currency-bearing cell: "1,234"  "-3,382"  "12744"  "4,975.50" */
  var BARE_NUM_RE = /^-?\d[\d,]*(?:\.\d+)?[kKmMbB]?$/;
  /* Skip script, style, and already-wrapped nodes. */
  var SKIP_TAGS = { SCRIPT:1, STYLE:1, NOSCRIPT:1, TEXTAREA:1, INPUT:1, CANVAS:1, SVG:1 };
  /* Classes that signal a cell holds a currency value — bare numbers inside these get blurred too. */
  var CURRENCY_CLASSES = ['num','kpi-value','portfolio-current','stat-value','value','cal-current'];

  function isInCurrencyCell(el){
    while (el && el !== document.body && el.classList) {
      for (var i = 0; i < CURRENCY_CLASSES.length; i++) {
        if (el.classList.contains(CURRENCY_CLASSES[i])) return true;
      }
      el = el.parentNode;
    }
    return false;
  }

  function wrapMoneyInNode(textNode){
    var parent = textNode.parentNode;
    if (!parent) return;
    var txt = textNode.nodeValue;
    if (!txt) return;
    /* Case 1: text contains an explicit currency symbol. */
    if (txt.indexOf('£') !== -1 || txt.indexOf('$') !== -1 || txt.indexOf('€') !== -1) {
      MONEY_RE.lastIndex = 0;
      if (MONEY_RE.test(txt)) {
        MONEY_RE.lastIndex = 0;
        var frag = document.createDocumentFragment();
        var lastIdx = 0;
        var m;
        while ((m = MONEY_RE.exec(txt)) !== null) {
          if (m.index > lastIdx) frag.appendChild(document.createTextNode(txt.slice(lastIdx, m.index)));
          var span = document.createElement('span');
          span.className = 'ed-money';
          span.textContent = m[0];
          frag.appendChild(span);
          lastIdx = m.index + m[0].length;
        }
        if (lastIdx < txt.length) frag.appendChild(document.createTextNode(txt.slice(lastIdx)));
        parent.replaceChild(frag, textNode);
        return;
      }
    }
    /* Case 2: bare number inside a currency-bearing cell (e.g., <td class="num">12,744</td>). */
    var trimmed = txt.trim();
    if (!trimmed) return;
    if (trimmed.indexOf('%') !== -1) return;      /* percentages stay visible */
    if (!BARE_NUM_RE.test(trimmed)) return;
    if (!isInCurrencyCell(parent)) return;
    /* Require 3+ digits — avoids blurring "3", "12", "2025" headers, years, counts. */
    var digitCount = trimmed.replace(/[^0-9]/g, '').length;
    if (digitCount < 3) return;
    /* Also skip 4-digit standalone years (1900–2099) that appear in .num columns. */
    if (/^(19|20)\d{2}$/.test(trimmed)) return;
    var start = txt.indexOf(trimmed);
    var fragB = document.createDocumentFragment();
    if (start > 0) fragB.appendChild(document.createTextNode(txt.slice(0, start)));
    var spanB = document.createElement('span');
    spanB.className = 'ed-money';
    spanB.textContent = trimmed;
    fragB.appendChild(spanB);
    var end = start + trimmed.length;
    if (end < txt.length) fragB.appendChild(document.createTextNode(txt.slice(end)));
    parent.replaceChild(fragB, textNode);
  }

  function walkAndWrap(root){
    if (!root) return;
    var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: function(n){
        if (!n.parentNode) return NodeFilter.FILTER_REJECT;
        if (SKIP_TAGS[n.parentNode.tagName]) return NodeFilter.FILTER_REJECT;
        if (n.parentNode.classList && n.parentNode.classList.contains('ed-money')) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    var nodes = [];
    var n;
    while ((n = walker.nextNode())) nodes.push(n);
    nodes.forEach(wrapMoneyInNode);
  }

  function applyPrivacy(on){
    document.body.classList.toggle('ed-privacy', on);
    var l = document.getElementById('ed-privacy-label');
    if (l) l.textContent = on ? 'Showing' : 'Private';
    var btn = document.getElementById('ed-privacy-toggle');
    if (btn) btn.classList.toggle('active', on);
  }

  function setupPrivacy(){
    var btn = document.getElementById('ed-privacy-toggle');
    if (!btn) return;
    walkAndWrap(document.body);
    /* Re-wrap when dashboards inject dynamic content (broker filters, chart tooltips). */
    var obs = new MutationObserver(function(muts){
      muts.forEach(function(m){
        m.addedNodes.forEach(function(node){
          if (node.nodeType === 1) walkAndWrap(node);
          else if (node.nodeType === 3) wrapMoneyInNode(node);
        });
      });
    });
    obs.observe(document.body, { childList: true, subtree: true });
    applyPrivacy(localStorage.getItem(PKEY) === '1');
    btn.addEventListener('click', function(){
      var next = !document.body.classList.contains('ed-privacy');
      localStorage.setItem(PKEY, next ? '1' : '0');
      applyPrivacy(next);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setupPrivacy);
  } else {
    setupPrivacy();
  }
})();
</script>
"""

CSS = r"""
:root {
  --bg: #f4f1ea;
  --bg-card: #fbf9f4;
  --bg-sunken: #ece8dd;
  --ink: #1a1a1a;
  --ink-2: #44433f;
  --ink-3: #7a7770;
  --ink-4: #a8a49a;
  --line: #dcd6c8;
  --line-2: #e6e1d3;
  --accent: #b8472c;
  --accent-2: #2d5b3e;
  --up: #2d5b3e;
  --down: #b8472c;
  --chip: #eae5d5;
  --f-display: "Fraunces", "Times New Roman", serif;
  --f-sans: "Inter", -apple-system, BlinkMacSystemFont, sans-serif;
  --f-mono: "JetBrains Mono", ui-monospace, monospace;
}
html[data-mood="dark"] {
  --bg: #0e0e10;
  --bg-card: #16161a;
  --bg-sunken: #0a0a0c;
  --ink: #f2ede2;
  --ink-2: #cdc8bd;
  --ink-3: #8a857a;
  --ink-4: #5a5650;
  --line: #26252a;
  --line-2: #1c1c20;
  --chip: #1e1d22;
  --accent: #e8a87c;
  --accent-2: #8cc29a;
  --up: #8cc29a;
  --down: #e8a87c;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: var(--f-sans);
  font-feature-settings: "ss01", "cv11";
  -webkit-font-smoothing: antialiased;
  background: var(--bg) !important;
  color: var(--ink) !important;
  padding: 28px 32px !important;
  min-height: 100vh;
  position: relative;
}
body::before {
  content: ""; position: fixed; inset: 0; pointer-events: none; z-index: 100;
  background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='160' height='160'><filter id='n'><feTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/><feColorMatrix values='0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 0.035 0'/></filter><rect width='100%25' height='100%25' filter='url(%23n)'/></svg>");
  opacity: .5; mix-blend-mode: multiply;
}
html[data-mood="dark"] body::before { mix-blend-mode: screen; opacity: .3; }

/* ── Editorial nav bar (shared across all dashboards) ──────── */
.ed-nav {
  display: flex; justify-content: space-between; align-items: center;
  gap: 16px; flex-wrap: wrap;
  padding: 0 0 16px 0;
  margin: 0 0 20px 0;
  border-bottom: 1px solid var(--ink);
}
.ed-nav-brand {
  font-family: var(--f-display); font-style: italic; font-weight: 500;
  font-size: 20px; letter-spacing: -0.01em; color: var(--ink);
  margin: 0;
}
.ed-nav-links { display: flex; gap: 4px; flex-wrap: wrap; }
.ed-nav-links a {
  font-family: var(--f-mono); font-size: 10px; letter-spacing: 0.1em;
  text-transform: uppercase; color: var(--ink-3) !important;
  padding: 4px 10px;
  border: 1px solid transparent; border-radius: 3px;
  text-decoration: none !important;
}
.ed-nav-links a:hover {
  color: var(--ink) !important; border-color: var(--line);
  background: var(--bg-card);
}
.ed-nav-links a.active {
  color: var(--ink) !important; border-color: var(--line);
  background: var(--bg-card);
}
.ed-theme-toggle, .ed-privacy-toggle {
  font-family: var(--f-mono); font-size: 11px; letter-spacing: 0.1em;
  text-transform: uppercase;
  background: var(--bg-card); color: var(--ink) !important;
  border: 1px solid var(--line); border-radius: 3px;
  padding: 5px 10px; cursor: pointer;
}
.ed-theme-toggle:hover, .ed-privacy-toggle:hover { border-color: var(--ink-3); }
.ed-nav-controls {
  display: flex; gap: 6px; align-items: center; flex-wrap: wrap;
}
.ed-privacy-toggle.active {
  background: var(--accent); color: var(--bg) !important;
  border-color: var(--accent);
}
/* Privacy mode: blur money values so a screen can be shown privately. */
body.ed-privacy .ed-money,
body.ed-privacy td.units-col,
body.ed-privacy th.units-col {
  filter: blur(6px);
  user-select: none;
  cursor: not-allowed;
  transition: filter 0.12s ease;
}
body:not(.ed-privacy) .ed-money,
body:not(.ed-privacy) td.units-col,
body:not(.ed-privacy) th.units-col {
  filter: none;
  transition: filter 0.12s ease;
}
::selection { background: var(--ink); color: var(--bg); }
button { font-family: inherit; }
a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }

/* Headings */
h1, .header h1 {
  font-family: var(--f-display) !important;
  font-weight: 500 !important;
  font-size: 38px !important;
  line-height: 1.1 !important;
  letter-spacing: -0.02em !important;
  color: var(--ink) !important;
  margin: 0 0 4px 0 !important;
  font-style: italic !important;
}
h2 {
  font-family: var(--f-display) !important;
  font-weight: 500 !important;
  font-size: 20px !important;
  letter-spacing: -0.01em !important;
  color: var(--ink) !important;
  margin: 0 0 12px 0 !important;
  font-style: normal !important;
}
h3 {
  font-family: var(--f-display) !important;
  font-weight: 500 !important;
  font-size: 16px !important;
  color: var(--ink) !important;
}

/* Subtitles and meta */
.sub, .subtitle {
  font-family: var(--f-mono) !important;
  font-size: 11px !important;
  color: var(--ink-3) !important;
  text-transform: uppercase !important;
  letter-spacing: 0.12em !important;
  margin: 0 0 24px 0 !important;
}

/* Header bar for eToro dashboard */
.header, .nav-bar {
  display: flex; justify-content: space-between; align-items: end;
  gap: 16px; flex-wrap: wrap;
  padding-bottom: 16px !important;
  border-bottom: 1px solid var(--ink) !important;
  margin-bottom: 20px !important;
}
.header-right, .nav-bar > * {
  font-family: var(--f-mono) !important;
  font-size: 11px !important;
  color: var(--ink-3) !important;
  text-transform: uppercase;
  letter-spacing: 0.1em;
}
.header-right a, .nav-bar a {
  color: var(--ink-3) !important; padding: 4px 8px;
  border: 1px solid transparent; border-radius: 3px;
}
.header-right a:hover, .nav-bar a:hover {
  color: var(--ink) !important; border-color: var(--line);
  background: var(--bg-card); text-decoration: none;
}

/* Cards and sections */
.card, .section {
  background: var(--bg-card) !important;
  border: 1px solid var(--line) !important;
  border-radius: 6px !important;
  padding: 20px !important;
  color: var(--ink) !important;
}
.card.grand {
  background: linear-gradient(135deg, var(--bg-card), var(--bg-sunken)) !important;
  border-color: var(--ink) !important;
}

/* KPI grid (eToro + macro) */
.kpi-grid, .grid-cards {
  display: grid;
  gap: 12px;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  margin-bottom: 20px;
}
@media (max-width: 900px) {
  .kpi-grid, .grid-cards {
    display: flex !important; overflow-x: auto; overflow-y: hidden;
    scroll-snap-type: x mandatory; -webkit-overflow-scrolling: touch;
    scrollbar-width: thin; padding-bottom: 4px;
  }
  .kpi-grid::-webkit-scrollbar, .grid-cards::-webkit-scrollbar { height: 3px; }
  .kpi-grid::-webkit-scrollbar-thumb, .grid-cards::-webkit-scrollbar-thumb { background: var(--line); }
  .kpi, .kpi-grid > .card, .grid-cards > .card {
    flex: 0 0 auto !important; min-width: 160px; scroll-snap-align: start;
  }
}
.kpi, .kpi-grid > .card, .grid-cards > .card {
  background: var(--bg-card) !important;
  border: 1px solid var(--line) !important;
  border-radius: 6px !important;
  padding: 14px 16px !important;
}
.kpi-label, .card .label {
  font-family: var(--f-mono) !important;
  font-size: 10px !important;
  color: var(--ink-3) !important;
  text-transform: uppercase !important;
  letter-spacing: 0.1em !important;
}
.kpi-value, .card .value {
  font-family: var(--f-display) !important;
  font-size: 28px !important;
  font-weight: 400 !important;
  color: var(--ink) !important;
  letter-spacing: -0.02em !important;
  margin-top: 4px !important;
  line-height: 1.1 !important;
}
.kpi-sub, .card .sub2 {
  font-family: var(--f-mono) !important;
  font-size: 11px !important;
  color: var(--ink-3) !important;
  margin-top: 4px !important;
}

/* Grid layouts */
.grid, .two-col, .three-col {
  display: grid;
  gap: 16px;
  margin-bottom: 16px;
}
.two-col { grid-template-columns: repeat(auto-fit, minmax(480px, 1fr)); }
.three-col { grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); }

/* Tables */
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td {
  padding: 8px 10px !important;
  text-align: left;
  border-bottom: 1px solid var(--line-2) !important;
  color: var(--ink) !important;
}
th {
  font-family: var(--f-mono) !important;
  color: var(--ink-3) !important;
  font-weight: 500 !important;
  font-size: 10px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.1em !important;
  background: var(--bg-sunken) !important;
}
td.num, th.num, .num {
  font-family: var(--f-mono) !important;
  font-variant-numeric: tabular-nums !important;
  text-align: right;
}
tbody tr:hover { background: var(--bg-sunken); }

/* Movement colors */
.pos { color: var(--up) !important; }
.neg { color: var(--down) !important; }
.delta { font-family: var(--f-mono); font-size: 12px; margin-left: 6px; }

/* Chips, badges, regions */
.chip, .region, .badge {
  display: inline-block;
  padding: 3px 9px !important;
  border-radius: 20px !important;
  font-family: var(--f-mono) !important;
  font-size: 10px !important;
  color: var(--ink-3) !important;
  background: var(--bg-sunken) !important;
  border: 1px solid var(--line) !important;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}
.broker-bar {
  display: flex; gap: 10px; margin-bottom: 14px; flex-wrap: wrap;
}

/* Stale/warning indicators */
.stale, .stale-warning, .warn {
  display: inline-block;
  padding: 2px 8px !important;
  border-radius: 3px !important;
  font-family: var(--f-mono) !important;
  font-size: 10px !important;
  background: #f5e4d7 !important;
  color: var(--accent) !important;
  letter-spacing: 0.05em;
  text-transform: uppercase;
  border: 1px solid var(--accent) !important;
}

/* Chart container */
.chart-wrap { background: var(--bg-card); padding: 12px; border-radius: 6px; }

/* Signal grid (eToro) */
.signal-grid {
  display: grid; gap: 10px;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
}

/* Footer */
.footer {
  font-family: var(--f-mono) !important;
  font-size: 10px !important;
  color: var(--ink-3) !important;
  text-transform: uppercase !important;
  letter-spacing: 0.1em !important;
  margin-top: 32px !important;
  padding-top: 16px !important;
  border-top: 1px solid var(--line) !important;
  text-align: center !important;
}

/* Responsive trims */
@media (max-width: 720px) {
  body { padding: 14px !important; }
  h1, .header h1 { font-size: 28px !important; }
  .kpi-value, .card .value { font-size: 22px !important; }
}
"""
