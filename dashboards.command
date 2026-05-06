#!/bin/bash
# ============================================================================
# dashboards.command — Mac equivalent of dashboards.bat
#
# Pulls latest code, copies fresh xlsx + JSON cache from Drive, refreshes all
# data feeds, regenerates every dashboard, copies static copies to Drive Upload.
#
# Run by double-clicking in Finder, or from Terminal:
#   ~/ClaudeCode/eToro/dashboards.command
# ============================================================================

set -e
cd "$(dirname "$0")"

source .venv/bin/activate

echo
echo "============================================================"
echo "  dashboards.command - $(date '+%Y-%m-%d %H:%M:%S')"
echo "============================================================"

mkdir -p data dashboards

echo
echo "[0/9] git pull (latest code from PC) ..."
git pull --ff-only || echo "  git pull skipped/failed - continuing with local code"

echo
echo "[0b/9] Refreshing eToro_Master data from Drive ..."
SYNC_DIR="$(python -c 'import sys; sys.path.insert(0,"scripts"); from paths import SYNC_DIR; print(SYNC_DIR)')"
for fname in eToro_Master.xlsx etoro_master.json bookmarks.json; do
  if [ -f "$SYNC_DIR/$fname" ]; then
    cp "$SYNC_DIR/$fname" "data/$fname"
    echo "  data/$fname"
  else
    echo "  WARNING: $SYNC_DIR/$fname not found - using existing data/$fname"
  fi
done

for fname in dashboard2.html health_dashboard.html; do
  if [ -f "$SYNC_DIR/$fname" ]; then
    cp "$SYNC_DIR/$fname" "dashboards/$fname"
    echo "  dashboards/$fname"
  else
    echo "  WARNING: $SYNC_DIR/$fname not found - dashboard2/health may be stale"
  fi
done

echo; echo "[1/9] Calendar sync ..."
python scripts/sync_calendar.py || echo "  WARNING: calendar sync failed"

echo; echo "[2/9] Gmail sync ..."
python scripts/sync_gmail.py || echo "  WARNING: gmail sync failed"

echo; echo "[3/9] Vault health notes sync ..."
python scripts/sync_vault.py || echo "  WARNING: vault sync failed"

echo; echo "[4/9] Earnings + dividends sync ..."
python scripts/sync_earnings_dividends.py || echo "  WARNING: earnings sync failed"

echo; echo "[5/9] Combined (T212) dashboard ..."
python run_combined.py || echo "  WARNING: T212/combined failed"

echo; echo "[6/9] Macro dashboard ..."
python run_macro.py || echo "  WARNING: macro failed"

echo; echo "[7/9] Stock exposure ..."
python scripts/sync_stock_exposure.py && \
python scripts/generate_exposure_dashboard.py || echo "  WARNING: exposure failed"

echo; echo "[8/9] Finances dashboard ..."
python scripts/sync_finances.py && \
python scripts/generate_finances_dashboard.py || echo "  WARNING: finances failed"

echo; echo "[8b/9] Regenerate eToro dashboard with live prices ..."
python scripts/generate_dashboard.py || echo "  WARNING: eToro dashboard regen failed"

echo; echo "[9/9] Build standalone dashboards + copy ALL to Drive Upload ..."
python scripts/build_static_dashboards.py || echo "  WARNING: static build failed"

echo
echo "============================================================"
echo "  dashboards.command - DONE"
echo "  Local: ~/ClaudeCode/eToro/dashboards/"
echo "  Drive: ~/.../My Drive/Upload/"
echo "============================================================"
