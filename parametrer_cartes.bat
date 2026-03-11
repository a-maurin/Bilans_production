@echo off
cd /d "%~dp0"

echo Ouverture de la configuration des profils de cartes...
call scripts\generateur_de_cartes\lancer_production_cartographique.bat --gui
pause
