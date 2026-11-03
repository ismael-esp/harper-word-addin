@echo off
rem Runs the Harper server hidden (no window) every time you log in to Windows.
set "VBS=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Harper for Word.vbs"
> "%VBS%" echo CreateObject("WScript.Shell").Run "cmd /c cd /d ""%~dp0"" && node server.mjs", 0, False
start "" wscript "%VBS%"
echo Done. Harper now starts hidden when Windows starts, and is running now.
echo To undo, run Uninstall.cmd or delete "Harper for Word.vbs" from shell:startup.
pause
