' Invisible wrapper for run_daily_briefing.bat
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c run_daily_briefing.bat", 0, False
