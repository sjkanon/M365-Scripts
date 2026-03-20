@ECHO OFF
SETLOCAL

:: ============================================================
:: start.bat — Windows setup helper (OOBE / USB)
:: Place on USB alongside GetAutoPilot.CMD and Get-WindowsAutoPilotInfo.ps1
::
:: OOBE usage:
::   1. Press Shift+F10 to open a command prompt
::   2. Find the USB drive letter (usually D: or E:)
::   3. Type:  D:\start.bat
:: ============================================================

:: Change working directory to the folder containing this script
cd /d "%~dp0"

:: Self-elevation — relaunch as Administrator if not already elevated
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:MENU
CLS
ECHO.
ECHO  ================================================
ECHO   Windows Setup Helper — USB Toolkit
ECHO  ================================================
ECHO.
ECHO   1  - Device Manager
ECHO   2  - Autopilot enrollment (save to CSV)
ECHO   3  - Remove hash file and re-run Autopilot (save to CSV)
ECHO   4  - Autopilot enrollment ONLINE (upload directly to Intune)
ECHO   5  - Windows Update (PSWindowsUpdate)
ECHO   6  - Install PowerShell 7
ECHO   7  - Enter product key
ECHO   8  - Join Active Directory domain
ECHO   9  - Restart
ECHO.
ECHO   A  - DO IT ALL (Autopilot online + Update + Restart)
ECHO   B  - Rename this device (NAME-SERIALNUMBER)
ECHO.
ECHO   0  - Exit
ECHO.
ECHO  ================================================
ECHO.
SET /P M=  Select option and press ENTER:

IF /I "%M%"=="1" GOTO DEVMGMT
IF /I "%M%"=="2" GOTO AUTOPILOT
IF /I "%M%"=="3" GOTO COMPHASH
IF /I "%M%"=="4" GOTO AUTOPILOT_ONLINE
IF /I "%M%"=="5" GOTO WINUPDATE
IF /I "%M%"=="6" GOTO INSTALLPS
IF /I "%M%"=="7" GOTO PRODUCTKEY
IF /I "%M%"=="8" GOTO ADJOIN
IF /I "%M%"=="9" GOTO RESTART
IF /I "%M%"=="A" GOTO DOITALL
IF /I "%M%"=="B" GOTO RENAMEPC
IF /I "%M%"=="0" GOTO EXIT

ECHO   Invalid option. Try again.
TIMEOUT /T 2 /NOBREAK >nul
GOTO MENU

:: ============================================================

:DEVMGMT
ECHO   Opening Device Manager...
start devmgmt.msc
GOTO MENU

:: ============================================================

:AUTOPILOT
ECHO   Starting Autopilot enrollment...
CALL "%~dp0GetAutoPilot.CMD"
GOTO MENU

:: ============================================================

:COMPHASH
ECHO   Removing existing hash file and re-running Autopilot...
IF EXIST "%~dp0compHash.csv" (
    del "%~dp0compHash.csv"
    ECHO   compHash.csv removed.
) ELSE (
    ECHO   No hash file found, continuing...
)
CALL "%~dp0GetAutoPilot.CMD"
GOTO MENU

:: ============================================================

:AUTOPILOT_ONLINE
ECHO.
ECHO   Autopilot enrollment — uploading directly to Intune...
ECHO   You will be prompted to sign in with your Microsoft 365 admin account.
ECHO.
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Get-WindowsAutoPilotInfo.ps1' -Online"
ECHO.
ECHO   Upload complete. Device should appear in Intune Autopilot within a few minutes.
ECHO.
PAUSE
GOTO MENU

:: ============================================================

:WINUPDATE
ECHO.
ECHO   Installing PSWindowsUpdate module and running updates...
ECHO   Requires internet connection.
ECHO.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Install-Module PSWindowsUpdate -Force -Scope CurrentUser; Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -AutoReboot"
ECHO.
ECHO   Windows Update complete. Device will reboot if updates were installed.
ECHO.
PAUSE
GOTO MENU

:: ============================================================

:INSTALLPS
ECHO.
ECHO   Installing PowerShell 7 via winget...
ECHO   Requires internet connection.
ECHO.
winget install --id Microsoft.PowerShell --source winget --silent --accept-source-agreements --accept-package-agreements
if %errorlevel% neq 0 (
    ECHO.
    ECHO   winget failed or not available on this device.
    ECHO   Alternative: download PowerShell 7 manually from https://aka.ms/powershell
    ECHO.
) ELSE (
    ECHO.
    ECHO   PowerShell 7 installed successfully.
    ECHO   Launch with: pwsh.exe
    ECHO.
)
PAUSE
GOTO MENU

:: ============================================================

:PRODUCTKEY
ECHO   Opening product key entry...
start slui.exe
GOTO MENU

:: ============================================================

:ADJOIN
ECHO.
ECHO   Join Active Directory domain
ECHO   You will be prompted for the domain name and admin credentials.
ECHO.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$domain = Read-Host 'Enter domain name (e.g. contoso.local)'; $cred = Get-Credential -Message ('Domain admin credentials for: ' + $domain); Add-Computer -DomainName $domain -Credential $cred -Force; Write-Host ''; Write-Host 'Domain join complete. Restart to apply.' -ForegroundColor Green"
ECHO.
PAUSE
GOTO MENU

:: ============================================================

:RESTART
ECHO   Restarting in 5 seconds... Press Ctrl+C to cancel.
shutdown -r -t 5
GOTO MENU

:: ============================================================

:DOITALL
CLS
ECHO.
ECHO  ================================================
ECHO   DO IT ALL
ECHO   Step 1: Autopilot enrollment (online - upload to Intune)
ECHO   Step 2: Windows Update
ECHO   Step 3: Restart
ECHO  ================================================
ECHO.

ECHO   [1/3] Removing old hash file...
IF EXIST "%~dp0compHash.csv" (
    del "%~dp0compHash.csv"
    ECHO         compHash.csv removed.
)

ECHO   [1/3] Running Autopilot enrollment (online - upload to Intune)...
ECHO         Sign in with your Microsoft 365 admin account when prompted.
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Get-WindowsAutoPilotInfo.ps1' -Online"

ECHO.
ECHO   [2/3] Installing PSWindowsUpdate and running updates...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Install-Module PSWindowsUpdate -Force -Scope CurrentUser; Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -AutoReboot"
ECHO         Windows Update complete.

ECHO.
ECHO   [3/3] All steps completed.
ECHO         The device will restart in 30 seconds.
ECHO         Press Ctrl+C to cancel the restart.
ECHO.
shutdown -r -t 30
PAUSE
GOTO MENU

:: ============================================================

:RENAMEPC
ECHO.
ECHO   Rename this device
ECHO   The serial number will be appended automatically.
ECHO   Format: NAME-SERIALNUMBER  (e.g. LAPTOP-ABC123)
ECHO.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$serial = (Get-WmiObject Win32_BIOS).SerialNumber.Trim(); $prefix = Read-Host 'Enter device name prefix (e.g. LAPTOP, DESKTOP, NB)'; $newName = ($prefix + '-' + $serial).ToUpper(); if ($newName.Length -gt 15) { Write-Host ('WARNING: name is ' + $newName.Length + ' characters — Windows allows max 15. Truncating.') -ForegroundColor Yellow; $newName = $newName.Substring(0,15) }; Rename-Computer -NewName $newName -Force; Write-Host (''); Write-Host ('Device renamed to: ' + $newName) -ForegroundColor Green; Write-Host 'Restart to apply the new name.'"
ECHO.
PAUSE
GOTO MENU

:: ============================================================

:EXIT
ENDLOCAL
EXIT
