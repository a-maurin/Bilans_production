@echo off
REM Lance l'interface de configuration des profils YAML sans console

cd /d "%~dp0"

REM Utilise pythonw (version sans console)
pythonw "tools\config_profils.py"