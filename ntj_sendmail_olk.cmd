@ECHO OFF
ECHO Launcher Python Script %1

REM ----- Configurazione -----
set "PYTHON=k:\tools\"

REM ----- Ricava nome script python con PATH completo -----
set "SCRIPT=%~dp0%~n0.py"

REM ----- Esecuzione -----
ECHO Sintassi filein.json
call "%PYTHON%pyn.cmd" "%SCRIPT%" %*