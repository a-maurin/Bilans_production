@echo off
setlocal
cd /d "%~dp0\..\.."

REM Wrapper Windows vers la CLI Python principale
python -m bilans %*

endlocal

