' Invisible wrapper for run_earnings_dividends_sync.bat
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c run_earnings_dividends_sync.bat", 0, False
