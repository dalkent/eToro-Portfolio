@echo off
cd /d C:\Users\Neil\ClaudeCode\eToro

:: Load environment variables from etoro.env
for /f "usebackq eol=# tokens=1,* delims==" %%A in ("etoro.env") do (
    if not "%%A"=="" if not "%%A"==" " set "%%A=%%B"
)

echo [%date% %time%] Running trade sync...

echo Running: sync_portfolio.py  (detects buys/sells, updates Portfolio / Closed Positions / Watchlist)
D:\Anaconda\python.exe scripts\sync_portfolio.py

echo [%date% %time%] Trade sync complete.
pause
