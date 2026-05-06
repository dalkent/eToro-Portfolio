@echo off
title Combined Dashboard Refresh (eToro + T212)
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

:: ── Logging ───────────────────────────────────────────────────────────────────
if not exist logs mkdir logs
set LOG=logs\combined_refresh.log
echo. >> %LOG%
echo ============================================================ >> %LOG%
echo  Combined Dashboard Refresh  [%date% %time%] >> %LOG%
echo ============================================================ >> %LOG%

:: Force Python into UTF-8 so unicode chars in print() don't crash under cp1252
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

:: run_combined.py loads etoro.env + t212.env itself, so no need to pre-load here
"%PYTHON%" run_combined.py >> %LOG% 2>&1
set RC=%ERRORLEVEL%
if %RC%==0 (
    echo Combined refresh OK [%date% %time%] >> %LOG%
) else (
    echo WARNING: run_combined.py exited %RC% >> %LOG%
)

:: ── Uptime Kuma heartbeat ────────────────────────────────────────────────────
:: Leave KUMA_URL blank (or the placeholder) to skip heartbeat. Fill it in after
:: creating a Push monitor in Kuma — see ..\Homelab\heartbeat.bat.snippet.
set KUMA_URL=http://localhost:3001/api/push/tw5fW0jp2e
if %RC%==0 (
    echo Calling Kuma [%date% %time%] status=up >> %LOG%
    curl.exe -v -m 10 --retry 3 "%KUMA_URL%?status=up&msg=OK" >> %LOG% 2>&1
) else (
    echo Calling Kuma [%date% %time%] status=down >> %LOG%
    curl.exe -v -m 10 --retry 3 "%KUMA_URL%?status=down&msg=ExitCode%RC%" >> %LOG% 2>&1
)
echo Kuma call done [%date% %time%] >> %LOG%

endlocal
