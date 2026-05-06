@echo off
title Macro Dashboard Refresh
setlocal

cd /d "%~dp0.."

:: ── Find Python ───────────────────────────────────────────────────────────────
set PYTHON=
if exist "D:\Anaconda\python.exe"              set PYTHON=D:\Anaconda\python.exe
if exist "C:\Anaconda3\python.exe"             set PYTHON=C:\Anaconda3\python.exe
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON=C:\ProgramData\Anaconda3\python.exe
if "%PYTHON%"=="" (
    where python >nul 2>&1
    if not errorlevel 1 (set PYTHON=python) else (
        echo ERROR: Python not found. & exit /b 1
    )
)

if not exist logs mkdir logs
set LOG=logs\macro_refresh.log
echo. >> %LOG%
echo ============================================================ >> %LOG%
echo  Macro Dashboard Refresh  [%date% %time%] >> %LOG%
echo ============================================================ >> %LOG%

set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

"%PYTHON%" run_macro.py >> %LOG% 2>&1
set RC=%ERRORLEVEL%
if %RC%==0 (
    echo Macro refresh OK [%date% %time%] >> %LOG%
) else (
    echo WARNING: run_macro.py exited %RC% >> %LOG%
)

:: Optional Kuma heartbeat — uncomment and fill in once you add a Push monitor.
:: set KUMA_URL=http://localhost:3001/api/push/TOKEN_HERE
:: if %RC%==0 (
::     curl.exe -fsS -m 10 --retry 3 "%KUMA_URL%?status=up&msg=OK" >nul 2>&1
:: ) else (
::     curl.exe -fsS -m 10 --retry 3 "%KUMA_URL%?status=down&msg=ExitCode%RC%" >nul 2>&1
:: )

endlocal
