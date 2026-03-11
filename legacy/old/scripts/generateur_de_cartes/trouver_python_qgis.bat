@echo off
setlocal EnableDelayedExpansion
REM Script de diagnostic : localiser python.exe associe a QGIS/OSGeo4W
REM Utilise si lancer_production_cartographique.bat echoue avec "QGIS Python introuvable"

echo Recherche de python.exe (QGIS / OSGeo4W)...
echo.

set "PF=C:\Program Files"
set "FOUND="

REM 1. Chemins connus
echo [Chemins connus]
if exist "%PF%\QGIS 3.40.15\bin\python.exe" (echo   OK: %PF%\QGIS 3.40.15\bin\python.exe & set "FOUND=1")
if exist "%PF%\QGIS 3.40.15\bin\python3.exe" (echo   OK: %PF%\QGIS 3.40.15\bin\python3.exe & set "FOUND=1")
if exist "%PF%\QGIS 3.40.15\osgeo4w\bin\python.exe" (echo   OK: %PF%\QGIS 3.40.15\osgeo4w\bin\python.exe & set "FOUND=1")
if exist "C:\OSGeo4W64\bin\python.exe" (echo   OK: C:\OSGeo4W64\bin\python.exe & set "FOUND=1")
if exist "C:\OSGeo4W\bin\python.exe" (echo   OK: C:\OSGeo4W\bin\python.exe & set "FOUND=1")

REM 2. Installation utilisateur OSGeo4W (AppData)
if exist "%LOCALAPPDATA%\Programs\OSGeo4W\bin\python.exe" (echo   OK: %LOCALAPPDATA%\Programs\OSGeo4W\bin\python.exe & set "FOUND=1")
if "!FOUND!"=="" (
    for /f "delims=" %%p in ('dir "%LOCALAPPDATA%\Programs\OSGeo4W\apps\Python*" /b /ad 2^>nul') do (
        if exist "%LOCALAPPDATA%\Programs\OSGeo4W\apps\%%p\python.exe" (echo   OK: %LOCALAPPDATA%\Programs\OSGeo4W\apps\%%p\python.exe & set "FOUND=1")
    )
)

REM 3. Recherche recursive PowerShell
echo.
echo [Recherche recursive dans C:\Program Files et AppData]
for /f "delims=" %%p in ('powershell -NoProfile -Command "Get-ChildItem -Path 'C:\Program Files', $env:LOCALAPPDATA -Filter 'python.exe' -Recurse -ErrorAction SilentlyContinue 2>$null | Where-Object { $_.FullName -match 'QGIS|OSGeo' } | ForEach-Object { $_.FullName }"') do (
    echo   Trouve: %%p
    set "FOUND=1"
)

if "!FOUND!"=="" (
    echo   Aucun python.exe trouve dans les dossiers QGIS/OSGeo
    echo.
    echo Si QGIS est installe ailleurs, indiquer le chemin dans lancer_production_cartographique.bat
)

echo.
pause
