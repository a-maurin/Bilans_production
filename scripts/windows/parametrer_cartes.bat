@echo off
setlocal
cd /d "%~dp0\..\.."

REM Ouverture de l'interface de configuration des cartes (QGIS)
call scripts\generateur_de_cartes\gui_config_cartes.bat %*

endlocal

