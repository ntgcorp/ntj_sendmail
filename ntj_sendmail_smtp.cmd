REM @echo off

REM ----- Configurazione -----
set "PYTHON=x pyn.cmd"

REM ----- Ricava nome script python con PATH completo -----
set "SCRIPT=%~dp0%~n0.py"

REM ----- Esecuzione -----
echo Sintassi filein.json
call "%PYTHON%" "%SCRIPT%" %*