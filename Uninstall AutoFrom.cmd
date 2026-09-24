@echo off
setlocal
echo AutoFrom uninstaller - close Outlook first.
echo.
powershell.exe -NoProfile -File "%~dp0uninstall.ps1" %*
set "AutoFromExitCode=%ERRORLEVEL%"
echo.
if not "%AutoFromExitCode%"=="0" echo Uninstall did not finish. Read the error above; no Outlook process was forcibly closed.
pause
exit /b %AutoFromExitCode%
