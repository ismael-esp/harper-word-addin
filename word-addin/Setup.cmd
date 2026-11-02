@echo off
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed. Install the LTS version from https://nodejs.org, then run Setup.cmd again.
  pause
  exit /b 1
)
echo Step 1 of 2: installing the local HTTPS certificate.
echo Windows will ask you to confirm. Click Yes.
call npx -y office-addin-dev-certs install
echo.
echo Step 2 of 2: registering the add-in with Word...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v HarperChecker /t REG_SZ /d "%~dp0manifest.xml" /f
echo.
echo Done. Close Word if it is open, run "Start Harper.cmd", then open Word.
pause
