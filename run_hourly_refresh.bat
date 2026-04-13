@echo off
title eToro Hourly Refresh
setlocal

cd /d "%~dp0"

:: ── Find Python ───────────────────────────────────────────────────────────────
set PYTHON=
if exist "D:\Anaconda\python.exe"         set PYTHON=D:\Anaconda\python.exe
if exist "C:\Anaconda3\python.exe"        set PYTHON=C:\Anaconda3\python.exe
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON=C:\ProgramData\Anaconda3\python.exe
if "%PYTHON%"=="" (
    where python >nul 2>&1
    if not errorlevel 1 (set PYTHON=python) else (
        echo ERROR: Python not found. & exit /b 1
    )
)

:: ── Load .env ─────────────────────────────────────────────────────────────────
if exist etoro.env (
    for /f "usebackq eol=# tokens=1,* delims==" %%A in ("etoro.env") do (
        if not "%%A"=="" if not "%%A"==" " set "%%A=%%B"
    )
)

:: ── Redirect all output to log ────────────────────────────────────────────────
set LOG=logs\hourly_refresh_run.log
echo. >> %LOG%
echo ============================================================ >> %LOG%
echo  eToro Hourly Refresh  [%date% %time%] >> %LOG%
echo ============================================================ >> %LOG%

:: ── 1. Valuations (run_daily.py) ──────────────────────────────────────────────
echo Running run_daily.py ... >> %LOG%
"%PYTHON%" run_daily.py >> %LOG% 2>&1
set RC1=%ERRORLEVEL%
if %RC1%==0 (
    echo run_daily.py OK >> %LOG%
) else (
    echo WARNING: run_daily.py exited %RC1% >> %LOG%
)

:: ── 2. Portfolio sync (sync_portfolio.py) — only if API keys present ──────────
if defined ETORO_PUBLIC_API_KEY (
    if defined ETORO_USER_KEY (
        echo Running sync_portfolio.py ... >> %LOG%
        "%PYTHON%" scripts\sync_portfolio.py >> %LOG% 2>&1
        set RC2=%ERRORLEVEL%
        if %RC2%==0 (
            echo sync_portfolio.py OK >> %LOG%
        ) else (
            echo WARNING: sync_portfolio.py exited %RC2% >> %LOG%
        )
    )
)

echo Hourly refresh complete [%date% %time%] >> %LOG%
endlocal
