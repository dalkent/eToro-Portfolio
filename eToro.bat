@echo off
REM ============================================================================
REM eToro.bat
REM
REM Refresh eToro portfolio data + recompute valuations.
REM Updates eToro_Master.xlsx + data\etoro_master.json + eToro_dashboard.html.
REM
REM Steps (delegated to run_all.py):
REM   1. sync_portfolio.py        — pull positions from eToro API
REM   2. valuation.py             — recompute DCF/DDM/EPV across all tickers
REM   3. generate_dashboard.py    — regenerate eToro_dashboard.html
REM   4. sync_xlsx_to_vault.py    — export xlsx -> data\etoro_master.json
REM
REM Use:
REM   eToro.bat                 — full run (~3-4 min)
REM   eToro.bat --no-sync       — skip eToro API, just revalue + regen dashboard
REM   eToro.bat --dash          — skip sync + valuations, just regen dashboard
REM ============================================================================

setlocal
cd /d "%~dp0"

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
echo   eToro.bat - data refresh + valuations
echo   %DATE% %TIME%
echo ============================================================
echo.

"%PYTHON%" run_all.py %*
exit /b %errorlevel%
