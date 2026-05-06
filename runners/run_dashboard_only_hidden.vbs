' Invisible wrapper for runners\run_dashboard_only.bat
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c run_dashboard_only.bat", 0, False
