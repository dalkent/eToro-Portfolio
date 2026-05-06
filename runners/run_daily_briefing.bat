@echo off
title eToro Daily Briefing
setlocal

cd /d "%~dp0.."

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
set LOG=logs\briefing_run.log
echo. >> %LOG%
echo ============================================================ >> %LOG%
echo  eToro Daily Briefing  [%date% %time%] >> %LOG%
echo ============================================================ >> %LOG%

:: ── Run tracker ───────────────────────────────────────────────────────────────
echo Running run_tracker.py ... >> %LOG%
"%PYTHON%" run_tracker.py >> %LOG% 2>&1
set RC=%ERRORLEVEL%

if %RC%==0 (
    echo Daily briefing COMPLETE [%date% %time%] >> %LOG%
) else (
    echo WARNING: run_tracker.py exited %RC% [%date% %time%] >> %LOG%
)

endlocal
