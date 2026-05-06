@echo off
REM ============================================================================
REM website.bat
REM
REM Rebuild + deploy daleyvaluations.com.
REM
REM Default behaviour:
REM   1. python scripts\build_site.py --refresh-prices
REM   2. git add . / commit / push
REM   3. Cloudflare auto-deploys within ~60 seconds
REM
REM Flags:
REM   --revalue   Re-run valuation.py + sync_xlsx_to_vault.py first.
REM               Use when the eToro workbook has changed and the JSON cache
REM               needs to be rebuilt before the site reads it.
REM
REM   --tracker   Also lock the FTSE Tracker .md to vault Drafts and render
REM               the 7 PNG images (hero + 6 tables). Use this on Tuesday
REM               mornings before drafting the Substack article.
REM
REM   --no-deploy Build site locally only - don't commit or push.
REM
REM Examples:
REM   website.bat                       — site refresh + deploy
REM   website.bat --revalue             — full pipeline (replaces full-refresh.bat)
REM   website.bat --tracker             — Tuesday Substack publish flow
REM   website.bat --revalue --tracker   — fresh data + site + tracker (full Tuesday)
REM ============================================================================

setlocal enabledelayedexpansion

set "ETORO_DIR=C:\Users\Neil\ClaudeCode\eToro"
set "SITE_DIR=C:\Users\Neil\ClaudeCode\daleyvaluations-site"

set REVALUE=0
set TRACKER=0
set NO_DEPLOY=0
:argloop
if "%~1"=="" goto argdone
if /I "%~1"=="--revalue"   set REVALUE=1
if /I "%~1"=="--tracker"   set TRACKER=1
if /I "%~1"=="--no-deploy" set NO_DEPLOY=1
shift
goto argloop
:argdone

set PYTHON=
if exist "D:\Anaconda\python.exe"               set PYTHON=D:\Anaconda\python.exe
if exist "C:\Anaconda3\python.exe"              set PYTHON=C:\Anaconda3\python.exe
if exist "C:\ProgramData\Anaconda3\python.exe"  set PYTHON=C:\ProgramData\Anaconda3\python.exe
if "%PYTHON%"=="" (
    where python >nul 2>&1
    if not errorlevel 1 (set PYTHON=python) else (
        echo ERROR: Python not found. & exit /b 1
    )
)

echo.
echo ============================================================
echo   website.bat - rebuild + deploy daleyvaluations.com
echo   revalue=%REVALUE%  tracker=%TRACKER%  no-deploy=%NO_DEPLOY%
echo   %DATE% %TIME%
echo ============================================================

REM ---- Step A: optional revalue ----
if %REVALUE%==1 (
    echo.
    echo [A] Re-running valuations + JSON export...
    pushd "%ETORO_DIR%"
    "%PYTHON%" scripts\valuation.py
    if errorlevel 1 (
        popd
        echo ERROR: valuation.py failed.
        exit /b 1
    )
    "%PYTHON%" scripts\sync_xlsx_to_vault.py
    if errorlevel 1 (
        popd
        echo ERROR: sync_xlsx_to_vault.py failed.
        exit /b 1
    )
    popd
)

REM ---- Step B: rebuild site ----
echo.
echo [B] Rebuilding daleyvaluations.com from JSON + fresh prices...
pushd "%SITE_DIR%"
"%PYTHON%" scripts\build_site.py --refresh-prices
if errorlevel 1 (
    popd
    echo ERROR: build_site.py failed.
    exit /b 1
)

REM ---- Step C: optional deploy ----
if %NO_DEPLOY%==1 (
    echo.
    echo [C] Skipping deploy ^(--no-deploy^). Site built locally only.
    popd
) else (
    echo.
    echo [C] Committing + pushing ^(Cloudflare auto-deploys^)...
    git add .
    git diff --cached --quiet
    if !errorlevel!==0 (
        echo   No site changes to commit.
    ) else (
        for /f %%t in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set DATESTAMP=%%t
        git commit -m "Refresh: !DATESTAMP!"
        if errorlevel 1 (
            popd
            echo ERROR: git commit failed.
            exit /b 1
        )
        git push
        if errorlevel 1 (
            popd
            echo ERROR: git push failed.
            exit /b 1
        )
    )
    popd
)

REM ---- Step D: optional tracker (Tuesday Substack flow) ----
if %TRACKER%==1 (
    echo.
    echo [D1] Locking FTSE Tracker .md to vault Drafts...
    pushd "%ETORO_DIR%"
    "%PYTHON%" scripts\generate_tracker.py
    if errorlevel 1 (
        echo WARNING: generate_tracker.py failed. Site is still live.
    )
    echo.
    echo [D2] Rendering 6 tracker-table PNGs...
    "%PYTHON%" scripts\generate_tracker_images.py
    if errorlevel 1 (
        echo WARNING: generate_tracker_images.py failed.
    )
    echo.
    echo [D3] Rendering hero cover image...
    "%PYTHON%" scripts\generate_tracker_hero.py
    if errorlevel 1 (
        echo WARNING: generate_tracker_hero.py failed.
    )
    popd
)

echo.
echo ============================================================
echo   website.bat - DONE
if %NO_DEPLOY%==1 (
    echo   Built locally; not pushed.
) else (
    echo   Live: https://daleyvaluations.com
)
if %TRACKER%==1 (
    echo   Tracker: ...\Daley's Brain\Projects\eToro ^& Investing\Drafts\
)
echo ============================================================
echo.

endlocal
exit /b 0
