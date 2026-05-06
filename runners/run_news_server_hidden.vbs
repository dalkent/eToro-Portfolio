' Invisible wrapper for run_news_server.bat — autostart mode (no browser)
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c set NEWS_NO_BROWSER=1 && run_news_server.bat", 0, False
