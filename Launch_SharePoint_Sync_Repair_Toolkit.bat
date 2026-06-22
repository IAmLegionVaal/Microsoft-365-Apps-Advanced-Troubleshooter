@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath '%~dp0SharePoint_Sync_Repair_Toolkit.ps1' -ErrorAction SilentlyContinue; & '%~dp0SharePoint_Sync_Repair_Toolkit.ps1'"
echo.
echo SharePoint sync repair workflow finished.
pause
endlocal
