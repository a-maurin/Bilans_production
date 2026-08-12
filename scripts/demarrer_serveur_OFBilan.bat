@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0.."
set "PROJECT_ROOT=%CD%"
set "PYTHONPATH=%PROJECT_ROOT%;%PROJECT_ROOT%\src;%PYTHONPATH%"

echo =====================================
echo     Lancement du serveur OFBilan
echo =====================================
echo.
echo Recherche de l'interpreteur Python de QGIS...

set QGIS_PYTHON=""

:: Cherche dans les installations standards de QGIS dans C:\Program Files\
for /d %%i in ("C:\Program Files\QGIS*") do (
    if exist "%%i\bin\python-qgis.bat" (
        set QGIS_PYTHON="%%i\bin\python-qgis.bat"
        goto :found
    )
    if exist "%%i\bin\python-qgis-ltr.bat" (
        set QGIS_PYTHON="%%i\bin\python-qgis-ltr.bat"
        goto :found
    )
)

:: Cherche dans les installations OSGeo4W
if exist "%LOCALAPPDATA%\Programs\OSGeo4W\bin\python-qgis.bat" (
    set QGIS_PYTHON="%LOCALAPPDATA%\Programs\OSGeo4W\bin\python-qgis.bat"
    goto :found
)
if exist "C:\OSGeo4W64\bin\python-qgis.bat" (
    set QGIS_PYTHON="C:\OSGeo4W64\bin\python-qgis.bat"
    goto :found
)
if exist "C:\OSGeo4W\bin\python-qgis.bat" (
    set QGIS_PYTHON="C:\OSGeo4W\bin\python-qgis.bat"
    goto :found
)

:found
if %QGIS_PYTHON%=="" (
    echo [ERREUR] Impossible de trouver une installation standard de QGIS.
    echo Veuillez modifier manuellement ce script pour pointer vers votre fichier "python-qgis.bat".
    pause
    exit /b 1
)

echo [OK] Interpreteur trouve : %QGIS_PYTHON%
echo.

:: Verification et installation des dependances Python requises
echo Verification des dependances Python...
cmd /c %QGIS_PYTHON% -c "import odf" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installation de odfpy (lecture fichiers ODS)...
    cmd /c %QGIS_PYTHON% -m pip install --quiet --user odfpy
    if errorlevel 1 (
        echo [ATTENTION] Impossible d'installer odfpy automatiquement.
        echo             Les fichiers PEJ/PA au format ODS ne pourront pas etre charges.
        echo             Installez manuellement : pip install odfpy
    ) else (
        echo [OK] odfpy installe avec succes.
    )
) else (
    echo [OK] odfpy disponible.
)
echo.
echo [OK] Demarrage du serveur...
echo.

cmd /c %QGIS_PYTHON% "%PROJECT_ROOT%\core\web\serveur.py"
if errorlevel 1 (
    echo.
    echo [ERREUR] Le serveur a rencontre une erreur au demarrage.
    pause
)
exit /b
