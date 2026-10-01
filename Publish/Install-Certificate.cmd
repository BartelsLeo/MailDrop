@echo off
rem Doppelklick-Start fuer Install-Certificate.ps1 (Doppelklick auf eine .ps1 oeffnet sie nur im Editor).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Certificate.ps1" %*
echo.
pause
