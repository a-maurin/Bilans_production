@echo off
REM Lance une interface graphique de lancement (PyQt) pour la production cartographique
REM en utilisant l'installation OSGeo4W utilisateur.

setlocal

REM Racine OSGeo4W de l'utilisateur (installation dans AppData\Local\Programs)
set "OSGEO4W_ROOT=%LOCALAPPDATA%\Programs\OSGeo4W"
if not exist "%OSGEO4W_ROOT%\bin\python.exe" (
    echo OSGeo4W introuvable dans %LOCALAPPDATA%\Programs\OSGeo4W
    echo Verifiez l'installation de QGIS / OSGeo4W.
    pause
    exit /b 1
)

REM Initialisation de l'environnement OSGeo4W / QGIS
REM Certaines installations OSGeo4W (notamment en AppData) n'ont pas qt5_env.bat / py3_env.bat.
REM On s'appuie sur python-qgis-ltr.bat qui prépare l'environnement QGIS/Python.

REM Se placer dans le dossier du script
pushd "%~dp0"

REM Lancer la GUI de lancement (scripts/generateur_de_cartes/gui_lancement_cartes.py)
call "%OSGEO4W_ROOT%\bin\python-qgis-ltr.bat" "gui_lancement_cartes.py"

popd
pause
exit /b %ERRORLEVEL%
