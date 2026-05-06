' Legacy pass-through — runs runners\run_combined_refresh.bat hidden.
' Kept at root because CombinedDashboard_15min scheduled task references this path.
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c runners\run_combined_refresh.bat", 0, False
