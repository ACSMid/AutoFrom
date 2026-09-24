@echo off
setlocal
powershell.exe -NoProfile -File "%~dp0install.ps1" %*
set "AutoFromExitCode=%ERRORLEVEL%"
echo.
pause
exit /b %AutoFromExitCode%
