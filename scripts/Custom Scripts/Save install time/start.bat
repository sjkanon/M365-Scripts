@ECHO OFF
CLS
:MENU
ECHO
ECHO Hallo Support,
ECHO Success met het installeren van de updates
ECHO
ECHO.
ECHO ...............................................
ECHO PRESS 1, 2, 3, 4, 5, 6 to select your task, or 10 to EXIT.
ECHO ...............................................
ECHO.
ECHO 1 - OPEN DEVMGMT
ECHO 2 - OPEN SETTINGS
ECHO 3 - START AUTOPILOT
ECHO 4 - REMOVE COMPHASH AND START AUTOPILOT
ECHO 5 - RESTART PC/LAPTOP
ECHO 6 - VUL PRO PRODUTKEY IN
ECHO 10 - EXIT CMD
ECHO.
SET /P M=Type 1, 2, 3, or 4 then press ENTER:
IF %M%==1 GOTO DEVMGMT
IF %M%==2 GOTO SETTINGS
IF %M%==3 GOTO AUTOPILOT
IF %M%==4 GOTO COMPHASH
IF %M%==5 GOTO RESTART
IF %M%==6 GOTO SLUI
IF %M%==10 GOTO exits
:DEVMGMT
start devmgmt.msc
GOTO MENU
:SETTINGS
start control.exe /name Microsoft.WindowsUpdate
GOTO MENU
:AUTOPILOT
start GetAutoPilot.CMD
GOTO MENU
:COMPHASH
del comphash.csv
start GetAutoPilot.CMD
GOTO MENU
:RESTART
shutdown -r -t 0
GOTO MENU
:SLUI
start slui.exe
GOTO MENU
:exits
EXIT
