@echo off
title Vault Health Sync
setlocal

cd /d "%~dp0.."

if not exist logs mkdir logs

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

"%PYTHON%" scripts\sync_vault.py >> logs\vault_sync.log 2>&1

endlocal
