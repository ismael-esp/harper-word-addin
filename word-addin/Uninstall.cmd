@echo off
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v HarperChecker /f
echo Harper removed from Word. To also remove the certificate, run: npx office-addin-dev-certs uninstall
pause
