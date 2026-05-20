"""Parse the Obsidian vault's Health notes and write data/health.json.

Reads:
  Personal/Health/Body Composition Log.md
  Personal/Health/Personal Bests.md
  Personal/Health/Food Log.md
  Personal/Health/Workouts/*.md

Writes:
  data/health.json

Env (in etoro.env):
  VAULT_ROOT — path to the vault root (default: C:\\Users\\Neil\\My Drive\\Daley's Brain)
"""
from __future__ import annotations

import json
import os
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from paths import VAULT_DIR as DEFAULT_VAULT

BASE = Path(__file__).parent.parent
DATA_DIR = BASE / "data"


def _load_env():
    env = BASE / "etoro.env"
    if not env.exists():
        return
    for line in env.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        os.environ.setdefault(k.strip(), v.strip())


def _strip_md(text: str) -> str:
    text = re.sub(r"\[\[([^\]|]+)(?:\|[^\]]+)?\]\]", r"\1", text)
    text = re.sub(r"[*_`]", "", text)
    return text.strip()


def _parse_table(lines: list[str], start: int) -> tuple[list[str], list[list[str]]]:
    """Parse a markdown table starting at `start`. Returns (headers, rows).

    Expects a header row, then a separator row of `---`, then data rows.
    Stops at the first blank line or non-pipe line.
    """
    if start >= len(lines):
        return [], []
    header = [c.strip() for c in lines[start].strip().strip("|").split("|")]
    if start + 1 >= len(lines) or "---" not in lines[start + 1]:
        return header, []
    rows = []
    for i in range(start + 2, len(lines)):
        line = lines[i].rstrip()
        if not line.strip():
            break
        if not line.lstrip().startswith("|"):
            break
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) == len(header):
            rows.append(cells)
    return header, rows


def _find_table(lines: list[str], after_header: str | None = None) -> int:
    """Return index of the first `|` table header row, optionally after a specific H2/H3."""
    i = 0
    if after_header:
        target = after_header.lower()
        for j, line in enumerate(lines):
            if line.strip().lower().startswith(target):
                i = j
                break
    for j in range(i, len(lines)):
        line = lines[j].strip()
        if line.startswith("|") and j + 1 < len(lines) and "---" in lines[j + 1]:
            return j
    return -1


# ── Body composition ─────────────────────────────────────────────────────────

def parse_body_comp(path: Path) -> dict:
    if not path.exists():
        return {}
    lines = path.read_text(encoding="utf-8").splitlines()
    # Find the "Full History" table — it has the complete series.
    idx = -1
    for i, line in enumerate(lines):
        if line.strip().lower().startswith("## full history") or line.strip().lower().startswith("## history"):
            idx = i
            break
    if idx == -1:
        idx = _find_table(lines, after_header="## summary")
    table_idx = _find_table(lines[idx:]) + idx if idx >= 0 else _find_table(lines)
    if table_idx < 0:
        return {}
    headers, rows = _parse_table(lines, table_idx)
    entries = []
    for row in rows:
        d = {}
        for h, v in zip(headers, row):
            d[h.lower().split(" ")[0]] = v
        # The date column may have a flag like "2026-01-05 ⚑" — strip the marker.
        date_str = re.sub(r"[^0-9\-]", "", d.get("date", "")).strip("-")[:10]
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            continue
        def _num(s):
            try:
                return float(re.sub(r"[^0-9.\-]", "", s))
            except (ValueError, TypeError):
                return None
        weight = _num(d.get("weight", ""))
        fat_kg = _num(d.get("fat", ""))
        muscle_kg = _num(d.get("muscle", ""))
        fat_pct = round(fat_kg / weight * 100, 1) if weight and fat_kg else None
        muscle_pct = round(muscle_kg / weight * 100, 1) if weight and muscle_kg else None
        entries.append({
            "date":       dt.isoformat(),
            "weight":     weight,
            "fat_kg":     fat_kg,
            "muscle_kg":  muscle_kg,
            "fat_pct":    fat_pct,
            "muscle_pct": muscle_pct,
            "visceral":   _num(d.get("visceral", "")),
            "met_age":    _num(d.get("met.", "") or d.get("met", "")),
            "bmi":        _num(d.get("bmi", "")),
        })
    entries.sort(key=lambda e: e["date"])
    if not entries:
        return {}
    latest = entries[-1]
    prev = entries[-2] if len(entries) >= 2 else None
    delta = {}
    if prev:
        for k in ("weight", "fat_kg", "muscle_kg", "fat_pct", "muscle_pct", "visceral", "bmi"):
            if latest.get(k) is not None and prev.get(k) is not None:
                delta[k] = round(latest[k] - prev[k], 2)
    return {
        "latest": latest,
        "previous": prev,
        "delta_vs_prev": delta,
        "trend": entries[-12:],  # last 12 entries for sparkline
        "count": len(entries),
    }


# ── Metabolic anchors (TDEE, BMR) from latest Health Report ─────────────────

def parse_latest_metabolic(reports_dir: Path) -> dict:
    """Walk Health Reports newest→oldest and return the first observed
    TDEE / BMR pair plus its source date.

    Targets the 'Additional data:' line in each report, e.g.:
        *Additional data: ... · BMR 1,645 kcal · TDEE 2,705 kcal · ...*

    Not every scan captures TDEE/BMR (some Boditrax sessions skip them), so
    the latest report may not have them — fall back through history.
    """
    if not reports_dir.exists() or not reports_dir.is_dir():
        return {}
    files = sorted(
        [p for p in reports_dir.iterdir() if p.suffix.lower() == ".md"],
        reverse=True,
    )
    tdee_rx = re.compile(r"TDEE\s+([0-9,]+)\s*kcal", re.IGNORECASE)
    bmr_rx  = re.compile(r"BMR\s+([0-9,]+)\s*kcal",  re.IGNORECASE)
    date_rx = re.compile(r"(\d{4}-\d{2}-\d{2})")
    out: dict = {}
    for fp in files:
        try:
            text = fp.read_text(encoding="utf-8", errors="replace")
        except Exception:
            continue
        if "tdee" not in text.lower():
            continue
        m_t = tdee_rx.search(text)
        if not m_t:
            continue
        m_b = bmr_rx.search(text)
        m_d = date_rx.search(fp.name) or date_rx.search(text[:300])
        try:
            tdee = int(m_t.group(1).replace(",", ""))
        except ValueError:
            continue
        out = {
            "tdee_kcal": tdee,
            "bmr_kcal":  int(m_b.group(1).replace(",", "")) if m_b else None,
            "source_date": m_d.group(1) if m_d else None,
            "source_file": fp.name,
        }
        break
    return out


# ── Daily activity (steps) ──────────────────────────────────────────────────

def parse_daily_activity(path: Path) -> dict:
    """Parse `Daily Activity Log.md` for per-day step counts.

    Reads every markdown table in the file (recent block, full history block,
    etc.) and merges rows on date. A row with `TBD` or blank steps is skipped.
    Estimated values (suffixed with `*`) are kept but flagged. Summary rows
    like '9-day total' / '9-day average' are filtered out by failing the date
    regex.

    Returns:
      {
        "latest": {"date": "...", "steps": N, "estimated": bool},
        "history": [{"date": "...", "steps": N, "estimated": bool}, ...],
        "rolling_7d_avg_latest": N,
        "rolling_7d": [{"date": "...", "value": N}, ...],  # parallel to history
        "all_time": {"count": N, "total": N, "avg": N},
        "goal_daily": 10000,
      }
    """
    if not path.exists():
        return {}
    lines = path.read_text(encoding="utf-8").splitlines()

    # Walk the file finding every table; collect step rows across all of them.
    by_date: dict[str, dict] = {}
    i = 0
    while i < len(lines):
        if lines[i].lstrip().startswith("|") and i + 1 < len(lines) and "---" in lines[i + 1]:
            headers, rows = _parse_table(lines, i)
            # Find the Date and Steps columns (case-insensitive, allow trailing words)
            h_lower = [h.lower() for h in headers]
            date_col  = next((k for k, h in enumerate(h_lower) if h.startswith("date")), None)
            steps_col = next((k for k, h in enumerate(h_lower) if "step" in h), None)
            if date_col is not None and steps_col is not None:
                for row in rows:
                    if len(row) <= max(date_col, steps_col):
                        continue
                    raw_date = row[date_col]
                    # Strip parenthetical day-of-week and bold markers: "2026-05-19 (Tue)"
                    m = re.search(r"\d{4}-\d{2}-\d{2}", raw_date)
                    if not m:
                        continue  # skips '**9-day total**' etc.
                    iso = m.group(0)
                    raw_steps = row[steps_col].strip()
                    if not raw_steps or raw_steps.upper() == "TBD":
                        continue
                    estimated = raw_steps.endswith("*")
                    digits = re.sub(r"[^0-9]", "", raw_steps)
                    if not digits:
                        continue
                    try:
                        steps = int(digits)
                    except ValueError:
                        continue
                    # Most-recent wins if the same date appears in multiple tables
                    # (recent block overrides full history).
                    by_date[iso] = {"date": iso, "steps": steps, "estimated": estimated}
            # Skip past this table
            j = i + 2
            while j < len(lines) and lines[j].lstrip().startswith("|"):
                j += 1
            i = j
        else:
            i += 1

    if not by_date:
        return {}

    history = sorted(by_date.values(), key=lambda e: e["date"])

    # 7-day trailing rolling average parallel to history (one entry per row).
    rolling: list[dict] = []
    from collections import deque
    window: deque[int] = deque(maxlen=7)
    for entry in history:
        window.append(entry["steps"])
        rolling.append({
            "date":  entry["date"],
            "value": round(sum(window) / len(window)),
        })

    total = sum(e["steps"] for e in history)
    return {
        "latest":                history[-1],
        "history":               history,
        "rolling_7d":            rolling,
        "rolling_7d_avg_latest": rolling[-1]["value"] if rolling else None,
        "all_time": {
            "count": len(history),
            "total": total,
            "avg":   round(total / len(history)),
        },
        "goal_daily": 10000,
    }


# ── Personal bests ───────────────────────────────────────────────────────────

def _pb_best_row(text: str) -> dict:
    # "55kg × 6 reps · 08/04/26"  →  {"best": "55kg × 6 reps", "date": "08/04/26"}
    m = re.match(r"^(.*?)\s*·\s*(\d{2}/\d{2}/\d{2})\s*$", text.strip())
    if m:
        return {"best": _strip_md(m.group(1)), "date": m.group(2)}
    return {"best": _strip_md(text), "date": ""}


def parse_personal_bests(path: Path) -> dict:
    if not path.exists():
        return {}
    lines = path.read_text(encoding="utf-8").splitlines()

    def _rows_under(heading: str) -> list[dict]:
        hl = heading.lower()
        start = -1
        for i, line in enumerate(lines):
            if line.strip().lower().startswith(hl.lower()):
                start = i
                break
        if start < 0:
            return []
        table_idx = _find_table(lines[start:])
        if table_idx < 0:
            return []
        headers, rows = _parse_table(lines, start + table_idx)
        out = []
        for row in rows:
            if len(row) < 2:
                continue
            rec = {"exercise": _strip_md(row[0]), "category": _strip_md(row[1]) if len(headers) >= 3 else ""}
            best_col = row[2] if len(row) >= 3 else row[1]
            rec.update(_pb_best_row(best_col))
            if len(row) >= 4:
                rec["standing"] = _strip_md(row[3])
            out.append(rec)
        return out

    # Categories tables (Push/Pull/Core/Cardio)
    categories = {}
    for cat in ("Push", "Pull", "Core", "Cardio"):
        hl = f"### {cat}"
        start = -1
        for i, line in enumerate(lines):
            if line.strip().lower() == hl.lower():
                start = i
                break
        if start < 0:
            continue
        table_idx = _find_table(lines[start:])
        if table_idx < 0:
            continue
        headers, rows = _parse_table(lines, start + table_idx)
        cat_rows = []
        for row in rows:
            if len(row) < 2:
                continue
            pb = _pb_best_row(row[1])
            cat_rows.append({"exercise": _strip_md(row[0]), **pb})
        if cat_rows:
            categories[cat.lower()] = cat_rows

    return {
        "recent":     _rows_under("**🏆 Recent PBs"),
        "stale":      _rows_under("**🎯 Targets to Beat"),
        "by_category": categories,
    }


# ── Workouts ────────────────────────────────────────────────────────────────

WORKOUT_FNAME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})\s+(.+?)\.md$")


def parse_workouts(dir_path: Path) -> dict:
    if not dir_path.exists():
        return {}
    files = []
    for f in dir_path.glob("*.md"):
        m = WORKOUT_FNAME_RE.match(f.name)
        if not m:
            continue
        try:
            d = datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            continue
        kind_raw = m.group(2)
        # Try frontmatter "workout:" first for canonical type.
        text = f.read_text(encoding="utf-8", errors="ignore")
        fm_match = re.search(r"^---\n(.*?)\n---", text, re.DOTALL)
        kind = ""
        if fm_match:
            for line in fm_match.group(1).splitlines():
                if line.strip().lower().startswith("workout:"):
                    kind = line.split(":", 1)[1].strip()
                    break
        if not kind:
            kind = kind_raw
        # Normalize e.g. "Back-and-biceps" → "Back & biceps"
        kind_clean = kind.replace("-", " ").strip()
        # Count body — exercise rows (table rows not header/separator)
        body_lines = text.splitlines()
        tbl_idx = _find_table(body_lines)
        exercises = set()
        set_count = 0
        if tbl_idx >= 0:
            headers, rows = _parse_table(body_lines, tbl_idx)
            for row in rows:
                if not row[0] or row[0].lower() in ("exercise", ""):
                    continue
                exercises.add(row[0].lower().strip())
                set_count += 1
        files.append({
            "date":      d.isoformat(),
            "kind":      kind_clean,
            "kind_raw":  kind,
            "file":      str(f),
            "exercises": len(exercises),
            "sets":      set_count,
        })
    files.sort(key=lambda x: x["date"], reverse=True)
    today = date.today()
    this_week_start = today - timedelta(days=today.weekday())
    this_week = [w for w in files if datetime.fromisoformat(w["date"]).date() >= this_week_start]

    # Simple next-workout prediction: recent Push/Pull/Core rotation.
    next_predicted = _predict_next(files)

    by_type: dict[str, int] = {}
    for w in this_week:
        k = w["kind"].split()[0]  # first word (Push/Pull/Core/Back/Chest)
        by_type[k] = by_type.get(k, 0) + 1

    return {
        "last":          files[0] if files else None,
        "this_week": {
            "count":   len(this_week),
            "by_type": by_type,
            "items":   this_week,
        },
        "next_predicted": next_predicted,
        "total_logged":   len(files),
    }


def _predict_next(files: list[dict]) -> dict:
    """Predict next workout based on last few entries."""
    if not files:
        return {"type": "Push", "reason": "starting fresh"}
    # Check last 3 to see pattern
    recent = [w["kind"].split()[0] for w in files[:3]]  # first word (Push/Pull/Core)
    # Use simple PPC rotation: Push → Pull → Core → Push
    rotation = ["Push", "Pull", "Core"]
    last = recent[0]
    if last in rotation:
        idx = rotation.index(last)
        nxt = rotation[(idx + 1) % len(rotation)]
        return {"type": nxt, "reason": f"after {last}"}
    # Fallback: suggest least-recent of the three
    counts = {t: 0 for t in rotation}
    for r in recent:
        if r in counts:
            counts[r] += 1
    nxt = min(counts, key=counts.get)
    return {"type": nxt, "reason": "least recent"}


# ── Food log ────────────────────────────────────────────────────────────────

FOOD_DAY_RE = re.compile(r"^###\s+(\d{4}-\d{2}-\d{2}).*$")
# Match optional ~, then digits possibly with commas and optional decimal.
_NUM_RE = re.compile(r"~?([0-9][0-9,]*(?:\.[0-9]+)?)")


def _extract_num(cell: str) -> float | None:
    m = _NUM_RE.search(cell or "")
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


def parse_food_log(path: Path, days: int = 7) -> dict:
    if not path.exists():
        return {}
    lines = path.read_text(encoding="utf-8").splitlines()
    # Split into day sections.
    days_data: dict[str, dict] = {}
    i = 0
    while i < len(lines):
        m = FOOD_DAY_RE.match(lines[i])
        if not m:
            i += 1
            continue
        day_str = m.group(1)
        # Collect until next ### or EOF
        section = []
        j = i + 1
        while j < len(lines) and not FOOD_DAY_RE.match(lines[j]):
            section.append(lines[j])
            j += 1
        # Extract last totals row (supports "Running total" and "Daily total").
        totals = None
        for line in section:
            stripped = line.strip()
            if not stripped.startswith("|"):
                continue
            low = stripped.lower()
            if "running total" not in low and "daily total" not in low:
                continue
            if True:
                cells = [c.strip() for c in stripped.strip("|").split("|")]
                # Typical: ["", "**Running total**", "**~1,672**", "**123g**", ...]
                nums = [_extract_num(c) for c in cells if c]
                nums = [n for n in nums if n is not None]
                if len(nums) >= 4:
                    totals = {
                        "calories":  int(nums[0]),
                        "protein_g": int(nums[1]),
                        "carbs_g":   int(nums[2]),
                        "fat_g":     int(nums[3]),
                    }
        # Workout day flag
        is_workout_day = any("🏋️" in line for line in [lines[i]] + section[:5])
        days_data[day_str] = {
            "date":    day_str,
            "totals":  totals,
            "workout_day": is_workout_day,
            "has_log": any(l.strip().startswith("|") for l in section),
        }
        i = j

    # Most recent N days with a totals row
    sorted_days = sorted(days_data.keys(), reverse=True)[:days]
    recent = [days_data[d] for d in sorted_days]
    today_key = date.today().isoformat()
    today_row = days_data.get(today_key, {"date": today_key, "totals": None, "has_log": False})
    return {
        "today":  today_row,
        "recent": recent,
        "targets": {
            "rest_day":    {"calories": 2100, "protein_g": 160, "carbs_g": 185, "fat_g": 65},
            "workout_day": {"calories": 2300, "protein_g": 160, "carbs_g": 215, "fat_g": 65},
        },
    }


# ── Main ────────────────────────────────────────────────────────────────────

def parse_goals(path: Path) -> dict:
    if not path.exists():
        return {}
    text = path.read_text(encoding="utf-8")
    # Isolate the Health & Fitness section
    m = re.search(r"##\s*Health.*?(?=\n##\s|\Z)", text, re.DOTALL | re.IGNORECASE)
    if not m:
        return {}
    section = m.group(0)

    def _find_row(label_re: str, row_source: str = section) -> list[str]:
        rx = re.compile(r"^\|\s*" + label_re + r".*$", re.IGNORECASE | re.MULTILINE)
        mm = rx.search(row_source)
        if not mm:
            return []
        return [c.strip() for c in mm.group(0).strip().strip("|").split("|")]

    def _num(s: str) -> float | None:
        m = re.search(r"-?\d+(?:\.\d+)?", s or "")
        return float(m.group(0)) if m else None

    # 90-day table has columns: Goal | Start | Target | Notes — Current is inline in Notes or Start.
    # Annual table has columns: Goal | Start | Target | Current | Status.
    # Parse both.
    ninety_re = re.search(r"###\s*90-Day Targets.*?(?=\n###\s|\Z)", section, re.DOTALL | re.IGNORECASE)
    annual_re = re.search(r"###\s*Annual.*?(?=\n###\s|\Z)", section, re.DOTALL | re.IGNORECASE)
    ninety = ninety_re.group(0) if ninety_re else ""
    annual = annual_re.group(0) if annual_re else ""

    out: dict[str, dict] = {}
    mappings = {
        "body_fat_pct": r"Body fat",
        "muscle_kg":    r"Muscle mass",
        "visceral":     r"Visceral fat",
        "training_per_week": r"Training sessions",
    }
    for key, label in mappings.items():
        ninety_row = _find_row(label, ninety)
        annual_row = _find_row(label, annual)
        entry: dict = {}
        if ninety_row and len(ninety_row) >= 3:
            entry["ninety_start"]  = _num(ninety_row[1])
            entry["ninety_target"] = _num(ninety_row[2])
        if annual_row and len(annual_row) >= 4:
            entry["annual_start"]   = _num(annual_row[1])
            # Annual "Target" cell may be a range like "16–17%" or "61kg"
            entry["annual_target_raw"] = annual_row[2]
            entry["annual_target"]     = _num(annual_row[2])
            entry["annual_current"]    = _num(annual_row[3])
        if entry:
            out[key] = entry
    return out


def parse_calisthenics_goals(path: Path) -> list[dict]:
    """Parse Personal/Health/Calisthenics Goals.md → list of
    {status, category, exercise, target, current, date, notes}.
    """
    if not path.exists():
        return []
    lines = path.read_text(encoding="utf-8").splitlines()
    idx = _find_table(lines)
    if idx < 0:
        return []
    headers, rows = _parse_table(lines, idx)
    # Normalise header keys (lowercase, first-word).
    keymap = [h.lower().split(" ")[0] for h in headers]
    # Expected: status | category | exercise | target | current | date | notes
    out = []
    for row in rows:
        rec = {k: v for k, v in zip(keymap, row)}
        # "Current Best" → stored under "current" (first-word keying), same for "Current"
        out.append({
            "status":   rec.get("status", "").strip(),
            "category": rec.get("category", "").strip(),
            "exercise": rec.get("exercise", "").strip(),
            "target":   rec.get("target", "").strip(),
            "current":  rec.get("current", "").strip() or "—",
            "date":     rec.get("date", "").strip(),
            "notes":    rec.get("notes", "").strip(),
        })
    return out


def parse_latest_health_report(reports_dir: Path) -> dict:
    """Find the newest `YYYY-MM-DD Health Report.md` and extract:
      - report_date: YYYY-MM-DD string
      - going_well: list of bullet strings
      - needs_attention: list of bullet strings
      - next_session: {title, rows: [{exercise, sets, target, notes}]}
      - headline: a short one-liner from the first summary paragraph
    """
    if not reports_dir.exists():
        return {}
    # Discover reports like 2026-04-23 Health Report.md
    candidates = []
    rx = re.compile(r"(\d{4}-\d{2}-\d{2}).*health report", re.IGNORECASE)
    for p in reports_dir.glob("*.md"):
        m = rx.search(p.name)
        if m:
            candidates.append((m.group(1), p))
    if not candidates:
        return {}
    candidates.sort(key=lambda x: x[0], reverse=True)
    report_date, path = candidates[0]
    text = path.read_text(encoding="utf-8")

    def _clean_bullet(line: str) -> str:
        s = line.strip()
        # Strip leading "- " or "* "
        if s.startswith(("- ", "* ")):
            s = s[2:]
        # Drop bold markers + inline code
        s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
        s = re.sub(r"`([^`]+)`", r"\1", s)
        return s.strip()

    # --- Summary section: "Going well" and "Needs attention" bullets ---
    summary_re = re.search(
        r"##\s*Summary:.*?(?=\n##\s|\Z)",
        text, re.DOTALL | re.IGNORECASE,
    )
    going_well: list[str] = []
    needs_attention: list[str] = []
    if summary_re:
        sec = summary_re.group(0)
        # Split on "Needs attention" header-ish line
        mm = re.search(r"\*\*Needs attention[:\*]*", sec, re.IGNORECASE)
        if mm:
            top = sec[: mm.start()]
            bot = sec[mm.start():]
        else:
            top, bot = sec, ""
        for line in top.splitlines():
            if line.lstrip().startswith(("- ", "* ")):
                going_well.append(_clean_bullet(line))
        for line in bot.splitlines():
            if line.lstrip().startswith(("- ", "* ")):
                needs_attention.append(_clean_bullet(line))

    # --- Next Session Suggestion: table rows ---
    next_session: dict = {}
    next_re = re.search(
        r"##\s*Next Session Suggestion[^\n]*\n(.*?)(?=\n##\s|\Z)",
        text, re.DOTALL | re.IGNORECASE,
    )
    if next_re:
        block = next_re.group(0)
        # Title from the heading
        title_m = re.search(r"##\s*Next Session Suggestion[:\s]*(.*)", block)
        title = (title_m.group(1).strip() if title_m else "Next Session").rstrip(":").strip()
        lines = block.splitlines()
        tidx = _find_table(lines)
        rows_out = []
        if tidx >= 0:
            headers, rows = _parse_table(lines, tidx)
            keymap = [h.lower().split(" ")[0] for h in headers]
            for row in rows:
                rec = {k: v for k, v in zip(keymap, row)}
                rows_out.append({
                    "exercise": rec.get("exercise", "").strip(),
                    "sets":     rec.get("sets", "").strip(),
                    "target":   rec.get("target", "").strip(),
                    "notes":    rec.get("notes", "").strip(),
                })
        next_session = {"title": title, "rows": rows_out}

    # --- Headline: first one-line takeaway-ish sentence from Summary ---
    headline = ""
    if going_well:
        headline = going_well[0]

    return {
        "report_date":     report_date,
        "report_file":     path.name,
        "headline":        headline,
        "going_well":      going_well,
        "needs_attention": needs_attention,
        "next_session":    next_session,
    }


def main() -> None:
    _load_env()
    vault_root = Path(os.environ.get("VAULT_ROOT") or str(DEFAULT_VAULT))
    health = vault_root / "Personal" / "Health"
    if not health.exists():
        sys.exit(f"Health folder not found at {health}")

    data = {
        "generated_at":   datetime.now().astimezone().isoformat(timespec="seconds"),
        "vault_root":     str(vault_root),
        "body_comp":      parse_body_comp(health / "Body Composition Log.md"),
        "personal_bests": parse_personal_bests(health / "Personal Bests.md"),
        "workouts":       parse_workouts(health / "Workouts"),
        "food":           parse_food_log(health / "Food Log.md"),
        "goals":          parse_goals(vault_root / "Goals.md"),
        "calisthenics":   parse_calisthenics_goals(health / "Calisthenics Goals.md"),
        "activity":       parse_daily_activity(health / "Daily Activity Log.md"),
        "metabolic":      parse_latest_metabolic(health / "Health Reports"),
        "latest_report":  parse_latest_health_report(health / "Health Reports"),
    }
    out = DATA_DIR / "health.json"
    out.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
