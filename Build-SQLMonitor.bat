@echo off
title SQL Server Monitor - Build Tool
color 0B
echo.
echo  ============================================================
echo   SQL Server Lite Monitor - EXE Builder
echo  ============================================================
echo.
echo  This script will compile SQLMonitor.ps1 into SQLMonitor.exe
echo  using PS2EXE (downloaded automatically via PowerShell).
echo.
echo  Requirements:
echo    - Windows 10 / Server 2016 or later
echo    - PowerShell 5.1+ (built into Windows)
echo    - Internet access (only needed ONCE to download PS2EXE)
echo.
echo  No Python, .NET SDK, or Visual Studio required!
echo.
pause

echo.
echo  [1/3] Checking PowerShell version...
powershell -NoProfile -Command "if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Host '  ERROR: PowerShell 5.1+ required.' -ForegroundColor Red; exit 1 } else { Write-Host ('  OK - PowerShell ' + $PSVersionTable.PSVersion.ToString()) -ForegroundColor Green }"
if %errorlevel% neq 0 goto :error

echo.
echo  [2/3] Installing PS2EXE module (if not already installed)...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "if (-not (Get-Module -ListAvailable -Name ps2exe)) { Write-Host '  Downloading PS2EXE...' -ForegroundColor Yellow; Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop; Write-Host '  PS2EXE installed.' -ForegroundColor Green } else { Write-Host '  PS2EXE already available.' -ForegroundColor Green }"
if %errorlevel% neq 0 (
    echo.
    echo  PS2EXE download failed. Trying alternative method...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "[Net.ServicePointManager]::SecurityProtocol='Tls12'; $wc=New-Object Net.WebClient; $wc.DownloadFile('https://github.com/MScholtes/PS2EXE/releases/latest/download/ps2exe.zip','%TEMP%\ps2exe.zip'); Expand-Archive '%TEMP%\ps2exe.zip' -DestinationPath '%TEMP%\ps2exe' -Force; Copy-Item '%TEMP%\ps2exe\*' -Destination '%USERPROFILE%\Documents\WindowsPowerShell\Modules\ps2exe' -Recurse -Force"
    if %errorlevel% neq 0 goto :error
)

echo.
echo  [3/3] Compiling SQLMonitor.ps1 into SQLMonitor.exe...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Import-Module ps2exe -Force; Invoke-ps2exe -inputFile '%~dp0SQLMonitor.ps1' -outputFile '%~dp0SQLMonitor.exe' -title 'SQL Server Lite Monitor' -description 'Lightweight SQL Server Monitoring Dashboard' -company 'SQL Monitor' -version '1.0.0.0' -iconFile '' -noConsole -requireAdmin -sta -verbose"

if %errorlevel% neq 0 goto :error

echo.
echo  ============================================================
echo   SUCCESS!  SQLMonitor.exe has been created.
echo  ============================================================
echo.
echo  You can now:
echo    1. Run SQLMonitor.exe directly on any Windows machine
echo    2. Copy SQLMonitor.exe to your SQL Servers - no install needed
echo    3. Right-click and "Run as Administrator" for best results
echo.
echo  The .exe works without PowerShell visible - it runs as a
echo  proper Windows desktop application.
echo.
start "" "%~dp0SQLMonitor.exe"
goto :end

:error
echo.
echo  ============================================================
echo   BUILD FAILED - See error above
echo  ============================================================
echo.
echo  If PS2EXE module install fails due to corporate policy,
echo  you can run the .ps1 directly instead:
echo.
echo    Right-click SQLMonitor.ps1 - Run with PowerShell
echo.
echo  Or unblock and run:
echo    powershell -ExecutionPolicy Bypass -File SQLMonitor.ps1
echo.

:end
pause
