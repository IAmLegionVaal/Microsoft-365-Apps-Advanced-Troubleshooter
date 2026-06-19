@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Microsoft 365 Apps Advanced Troubleshooter - Created by Dewald Pretorius

rem Remove the downloaded-file marker from this launcher and the PowerShell script.
rem Windows may still display one warning before this BAT is allowed to run the first time.
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command ^
  "$targets = @('%~f0', '%~dp0Microsoft_365_Apps_Advanced_Troubleshooter.ps1'); foreach ($target in $targets) { if (Test-Path -LiteralPath $target) { try { Unblock-File -LiteralPath $target -ErrorAction SilentlyContinue } catch {}; try { Remove-Item -LiteralPath $target -Stream Zone.Identifier -ErrorAction SilentlyContinue } catch {} } }" >nul 2>&1

echo ============================================================
echo   Microsoft 365 Apps Advanced Troubleshooter
echo   Created by Dewald Pretorius
echo ============================================================
echo.
echo Starting in standard user mode.
echo Some advanced repairs may require Run as administrator.
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Microsoft_365_Apps_Advanced_Troubleshooter.ps1"

if errorlevel 1 (
    echo.
    echo The troubleshooter exited with an error.
)

echo.
pause
endlocal
