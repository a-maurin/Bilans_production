@echo off
setlocal EnableDelayedExpansion
REM Lancer la production cartographique avec le Python de QGIS
REM Détecte automatiquement la version la plus récente de QGIS installée

set "QGIS_PYTHON="
set "PF=C:\Program Files"

REM 0. Priorite : OSGeo4W AppData - si --gui deleguer a lancer_osgeo4w.bat ; sinon generer les cartes avec le meme Python
set "LOCALAPPDATA_OSGEO=%LOCALAPPDATA%\Programs\OSGeo4W"
if exist "%LOCALAPPDATA_OSGEO%\bin\python.exe" (
    if "%~1"=="--gui" (
        call "%~dp0lancer_osgeo4w.bat" %*
        exit /b %ERRORLEVEL%
    )
    REM Generation de cartes : lancer production_cartographique.py avec les arguments (periode, dept, profil)
    set "QGIS_PYTHON=%LOCALAPPDATA_OSGEO%\bin\python.exe"
    for %%A in ("!QGIS_PYTHON!") do set "OSGEO4W_BIN=%%~dpA"
    set "OSGEO4W_ROOT=!OSGEO4W_BIN:~0,-4!"
    call "!OSGEO4W_ROOT!\bin\o4w_env.bat"
    set "PYTHONHOME=!OSGEO4W_ROOT!\apps\Python312"
    path "!OSGEO4W_ROOT!\apps\qgis-ltr\bin";!PATH!
    set "QGIS_PREFIX_PATH=!OSGEO4W_ROOT:/=!/apps/qgis-ltr"
    set "GDAL_FILENAME_IS_UTF8=YES"
    set "QT_QPA_PLATFORM=offscreen"
    set "PYTHONPATH=!OSGEO4W_ROOT!\apps\qgis-ltr\python;!PYTHONPATH!"
    cd /d "%~dp0..\..\.."
    set "SCRIPT_DIR=%~dp0"
    "!QGIS_PYTHON!" "!SCRIPT_DIR!production_cartographique.py" %*
    if errorlevel 1 (
        echo.
        echo Le script a echoue. Verifier les messages ci-dessus.
    )
    pause
    exit /b %ERRORLEVEL%
)

REM 1. Priorite : QGIS 3.40.15 (version la plus recente connue)
if exist "%PF%\QGIS 3.40.15\bin\python.exe" set "QGIS_PYTHON=%PF%\QGIS 3.40.15\bin\python.exe"
if "!QGIS_PYTHON!"=="" if exist "%PF%\QGIS 3.40.15\bin\python3.exe" set "QGIS_PYTHON=%PF%\QGIS 3.40.15\bin\python3.exe"
if "!QGIS_PYTHON!"=="" (
    for /f "delims=" %%p in ('dir "%PF%\QGIS 3.40.15\apps\Python*" /b /ad 2^>nul') do (
        if exist "%PF%\QGIS 3.40.15\apps\%%p\python.exe" set "QGIS_PYTHON=%PF%\QGIS 3.40.15\apps\%%p\python.exe"
    )
)
REM Structure OSGeo4W imbriquee
if "!QGIS_PYTHON!"=="" if exist "%PF%\QGIS 3.40.15\osgeo4w\bin\python.exe" set "QGIS_PYTHON=%PF%\QGIS 3.40.15\osgeo4w\bin\python.exe"
if "!QGIS_PYTHON!"=="" (
    for /f "delims=" %%p in ('dir "%PF%\QGIS 3.40.15\osgeo4w\apps\Python*" /b /ad 2^>nul') do (
        if exist "%PF%\QGIS 3.40.15\osgeo4w\apps\%%p\python.exe" set "QGIS_PYTHON=%PF%\QGIS 3.40.15\osgeo4w\apps\%%p\python.exe"
    )
)

REM 2. Parcourir les autres versions QGIS (si 3.40.15 non trouve)
if "!QGIS_PYTHON!"=="" for /f "delims=" %%d in ('dir "%PF%\QGIS*" /b /ad /o-n 2^>nul') do (
    if "!QGIS_PYTHON!"=="" (
        if exist "%PF%\%%d\bin\python.exe" set "QGIS_PYTHON=%PF%\%%d\bin\python.exe"
    )
    if "!QGIS_PYTHON!"=="" (
        for /f "delims=" %%p in ('dir "%PF%\%%d\apps\Python*" /b /ad 2^>nul') do (
            if exist "%PF%\%%d\apps\%%p\python.exe" set "QGIS_PYTHON=%PF%\%%d\apps\%%p\python.exe"
        )
    )
)

REM 3. Fallback : OSGeo4W (installation classique)
if "!QGIS_PYTHON!"=="" if exist "C:\OSGeo4W64\bin\python.exe" set "QGIS_PYTHON=C:\OSGeo4W64\bin\python.exe"
if "!QGIS_PYTHON!"=="" if exist "C:\OSGeo4W\bin\python.exe" set "QGIS_PYTHON=C:\OSGeo4W\bin\python.exe"

REM 4. Recherche recursive PowerShell (installations atypiques)
if "!QGIS_PYTHON!"=="" (
    for /f "delims=" %%p in ('powershell -NoProfile -Command "Get-ChildItem -Path 'C:\Program Files', $env:LOCALAPPDATA -Filter 'python.exe' -Recurse -ErrorAction SilentlyContinue 2>$null | Where-Object { $_.FullName -match 'QGIS|OSGeo' } | Select-Object -First 1 -ExpandProperty FullName"') do (
        if exist "%%p" set "QGIS_PYTHON=%%p"
    )
)

REM 5. Installation utilisateur OSGeo4W AppData (deja traite en 0 si present)
if "!QGIS_PYTHON!"=="" (
    for /f "delims=" %%p in ('dir "%LOCALAPPDATA_OSGEO%\apps\Python*" /b /ad 2^>nul') do (
        if exist "%LOCALAPPDATA_OSGEO%\apps\%%p\python.exe" set "QGIS_PYTHON=%LOCALAPPDATA_OSGEO%\apps\%%p\python.exe"
    )
)
REM 5b. Recherche recursive dans AppData
if "!QGIS_PYTHON!"=="" (
    for /f "delims=" %%p in ('powershell -NoProfile -Command "$p=[Environment]::GetFolderPath('LocalApplicationData'); Get-ChildItem -Path $p -Filter 'python.exe' -Recurse -ErrorAction SilentlyContinue 2>$null | Where-Object { $_.FullName -match 'QGIS|OSGeo' } | Select-Object -First 1 -ExpandProperty FullName"') do (
        if exist "%%p" set "QGIS_PYTHON=%%p"
    )
)

REM 6. Fichier de configuration manuel (si detection auto echoue)
if "!QGIS_PYTHON!"=="" (
    if exist "%~dp0qgis_python_path.txt" (
        set /p QGIS_PYTHON=<"%~dp0qgis_python_path.txt"
        for /f "tokens=*" %%a in ("!QGIS_PYTHON!") do set "QGIS_PYTHON=%%~a"
        if not exist "!QGIS_PYTHON!" set "QGIS_PYTHON="
    )
)

if "!QGIS_PYTHON!"=="" (
    echo QGIS Python introuvable.
    echo.
    echo Solutions :
    echo   1. Executer trouver_python_qgis.bat pour diagnostiquer
    echo   2. Creer le fichier qgis_python_path.txt dans ce dossier
    echo      avec une seule ligne : le chemin complet vers python.exe
    echo      Exemple : C:\Program Files\QGIS 3.40.15\bin\python.exe
    echo.
    echo Pour trouver python.exe : ouvrir QGIS, menu Extensions -^> Console Python,
    echo puis taper : import sys ; print(sys.executable)
    pause
    exit /b 1
)

cd /d "%~dp0..\..\.."
set "SCRIPT_DIR=%~dp0"
REM Configuration environnement pour OSGeo4W : reprend le setup de qgis_process-qgis-ltr.bat
echo !QGIS_PYTHON! | findstr /i "OSGeo4W" >nul
if not errorlevel 1 (
    for %%A in ("!QGIS_PYTHON!") do set "OSGEO4W_BIN=%%~dpA"
    set "OSGEO4W_ROOT=!OSGEO4W_BIN:~0,-4!"
    call "!OSGEO4W_ROOT!\bin\o4w_env.bat"
    set "PYTHONHOME=!OSGEO4W_ROOT!\apps\Python312"
    path "!OSGEO4W_ROOT!\apps\qgis-ltr\bin";!PATH!
    set "QGIS_PREFIX_PATH=!OSGEO4W_ROOT:/=!/apps/qgis-ltr"
    set "GDAL_FILENAME_IS_UTF8=YES"
    set "QT_QPA_PLATFORM=offscreen"
    set "PYTHONPATH=!OSGEO4W_ROOT!\apps\qgis-ltr\python;!PYTHONPATH!"
)
"%QGIS_PYTHON%" "%SCRIPT_DIR%production_cartographique.py" %*
if errorlevel 1 (
    echo.
    echo Le script a echoue. Verifier les messages ci-dessus.
)
pause
