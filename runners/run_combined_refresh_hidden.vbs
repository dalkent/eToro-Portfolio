' Invisible wrapper for run_combined_refresh.bat
' Task Scheduler points at this .vbs instead of the .bat — no console window flashes.
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c run_combined_refresh.bat", 0, False
