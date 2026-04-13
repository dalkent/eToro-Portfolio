@echo off
title eToro Daily Valuation Run
setlocal

:: ── Working directory: same folder as this bat file ──────────────────────────
cd /d "%~dp0"

:: ── Find Python (try Anaconda D:\, then C:\, then system PATH) ───────────────
set PYTHON=
if exist "D:\Anaconda\python.exe"  set PYTHON=D:\Anaconda\python.exe
if exist "C:\Anaconda3\python.exe" set PYTHON=C:\Anaconda3\python.exe
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON=C:\ProgramData\Anaconda3\python.exe
if "%PYTHON%"=="" (
    where python >nul 2>&1
    if not errorlevel 1 (
        set PYTHON=python
    ) else (
        echo.
        echo ERROR: Python not found.
        echo   Checked: D:\Anaconda, C:\Anaconda3, C:\ProgramData\Anaconda3, and PATH
        echo   Please update the PYTHON path in this bat file.
        echo.
        pause
        exit /b 1
    )
)

:: ── Load environment variables from etoro.env ─────────────────────────────────
if exist etoro.env (
    for /f "usebackq eol=# tokens=1,* delims==" %%A in ("etoro.env") do (
        if not "%%A"=="" if not "%%A"==" " set "%%A=%%B"
    )
)

:: ── Run ───────────────────────────────────────────────────────────────────────
echo.
echo ============================================================
echo  eToro Daily Valuation  [%date% %time%]
echo  Python: %PYTHON%
echo  Folder: %CD%
echo ============================================================
echo.

echo Running: valuation.py  (Yahoo Finance valuations + BTC price)
echo This takes 2-3 minutes for 134 tickers — please wait...
echo.

"%PYTHON%" scripts\valuation.py
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE%==0 (
    echo ============================================================
    echo  Valuation run COMPLETE  [%date% %time%]
    echo  Check logs\valuation_run.log for the full output.
    echo ============================================================
) else (
    echo ============================================================
    echo  WARNING: valuation.py exited with code %EXIT_CODE%
    echo  Check logs\valuation_run.log for error details.
    echo ============================================================
)

echo.
echo Running: generate_dashboard.py  (Portfolio dashboard)
echo.

"%PYTHON%" scripts\generate_dashboard.py
set DASH_CODE=%ERRORLEVEL%

echo.
if %DASH_CODE%==0 (
    echo ============================================================
    echo  Dashboard generated: dashboard.html
    echo  Open in browser to view your portfolio.
    echo ============================================================
) else (
    echo ============================================================
    echo  WARNING: generate_dashboard.py exited with code %DASH_CODE%
    echo ============================================================
)

echo.
if %EXIT_CODE%==0 if %DASH_CODE%==0 (
    echo  Daily run COMPLETE  [%date% %time%]
) else (
    echo  Daily run finished with warnings — check output above.
)

echo.
pause
endlocal
