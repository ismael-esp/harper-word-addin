@echo off
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v HarperChecker /f
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Harper for Word.vbs" 2>nul
echo Harper removed from Word. To also remove the certificate, run: npx office-addin-dev-certs uninstall
pause
