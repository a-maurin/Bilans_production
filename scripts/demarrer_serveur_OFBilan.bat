@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d "%~dp0.."
set "PROJECT_ROOT=%CD%"

echo =====================================
echo     Lancement du serveur OFBilan
echo =====================================
echo.
echo Recherche de l'interpreteur Python de QGIS...

set "QGIS_PYTHON="

:: Priorite : wrappers python-qgis*.bat qui configurent PYTHONHOME correctement
for /d %%i in ("C:\Program Files\QGIS*") do (
    if "!QGIS_PYTHON!"=="" if exist "%%i\bin\python-qgis-ltr.bat" set "QGIS_PYTHON=%%i\bin\python-qgis-ltr.bat"
)
for /d %%i in ("C:\Program Files\QGIS*") do (
    if "!QGIS_PYTHON!"=="" if exist "%%i\bin\python-qgis.bat" set "QGIS_PYTHON=%%i\bin\python-qgis.bat"
)
if "!QGIS_PYTHON!"=="" if exist "%LOCALAPPDATA%\Programs\OSGeo4W\bin\python-qgis-ltr.bat" set "QGIS_PYTHON=%LOCALAPPDATA%\Programs\OSGeo4W\bin\python-qgis-ltr.bat"
if "!QGIS_PYTHON!"=="" if exist "%LOCALAPPDATA%\Programs\OSGeo4W\bin\python-qgis.bat" set "QGIS_PYTHON=%LOCALAPPDATA%\Programs\OSGeo4W\bin\python-qgis.bat"
if "!QGIS_PYTHON!"=="" if exist "C:\OSGeo4W64\bin\python-qgis-ltr.bat" set "QGIS_PYTHON=C:\OSGeo4W64\bin\python-qgis-ltr.bat"

if "!QGIS_PYTHON!"=="" (
    echo [ERREUR] Impossible de trouver python-qgis-ltr.bat ou python-qgis.bat.
    echo         Verifiez que QGIS est installe dans C:\Program Files.
    pause
    exit /b 1
)

echo [OK] Interpreteur trouve : "!QGIS_PYTHON!"
echo.

:: Verifier et installer odfpy via un script Python temporaire
:: (evite les problemes de redirection sur les wrappers .bat)
if "%DEBUG%"=="1" echo Verification de la bibliotheque odfpy...
if "%OFBILAN_DEBUG%"=="1" echo Verification de la bibliotheque odfpy...
set "TMP_CHECK=%TEMP%\ofbilan_odf_check.py"
(
    echo import sys
    echo try:
    echo     import odf
    echo except ImportError:
    echo     import subprocess
    echo     subprocess.check_call^([sys.executable, '-m', 'pip', 'install', '--quiet', '--user', 'odfpy']^)
    echo     print^('[OK] odfpy installe avec succes.'^)
) > "%TMP_CHECK%"

call "!QGIS_PYTHON!" "%TMP_CHECK%"
del "%TMP_CHECK%" >nul 2>&1

echo.
echo [OK] Demarrage du serveur...
echo.

set "DEBUG_ARG="
if "%DEBUG%"=="1" set "DEBUG_ARG=--debug"
if "%OFBILAN_DEBUG%"=="1" set "DEBUG_ARG=--debug"

:: NE PAS definir PYTHONPATH ici : serveur.py gere son propre sys.path
:: et python-qgis*.bat configure PYTHONHOME correctement
call "!QGIS_PYTHON!" "%PROJECT_ROOT%\core\web\serveur.py" !DEBUG_ARG!
if errorlevel 1 (
    echo.
    echo [ERREUR] Le serveur s'est arrete avec une erreur.
    pause
)
