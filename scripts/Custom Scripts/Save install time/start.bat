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
:: (ensures GetAutoPilot.CMD is found regardless of USB drive letter)
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
ECHO   2  - Autopilot enrollment
ECHO   3  - Remove hash file and re-run Autopilot
ECHO   4  - Enter product key
ECHO   5  - Restart
ECHO   0  - Exit
ECHO.
ECHO  ================================================
ECHO.
SET /P M=  Select option and press ENTER:

IF "%M%"=="1" GOTO DEVMGMT
IF "%M%"=="2" GOTO AUTOPILOT
IF "%M%"=="3" GOTO COMPHASH
IF "%M%"=="4" GOTO PRODUCTKEY
IF "%M%"=="5" GOTO RESTART
IF "%M%"=="0" GOTO EXIT

ECHO   Invalid option. Try again.
TIMEOUT /T 2 /NOBREAK >nul
GOTO MENU

:DEVMGMT
ECHO   Opening Device Manager...
start devmgmt.msc
GOTO MENU

:AUTOPILOT
ECHO   Starting Autopilot enrollment...
CALL "%~dp0GetAutoPilot.CMD"
GOTO MENU

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

:PRODUCTKEY
ECHO   Opening product key entry...
start slui.exe
GOTO MENU

:RESTART
ECHO   Restarting in 5 seconds... Press Ctrl+C to cancel.
shutdown -r -t 5
GOTO MENU

:EXIT
ENDLOCAL
EXIT
