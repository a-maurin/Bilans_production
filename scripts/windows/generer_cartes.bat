@echo off
setlocal
cd /d "%~dp0\..\.."

REM Wrapper Windows pour la génération de cartes
python scripts\generer_cartes.py %*

endlocal

