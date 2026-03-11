@echo off
REM Test rapide import PyQGIS (supprimer après vérification)
set "OSGEO4W_ROOT=%LOCALAPPDATA%\Programs\OSGeo4W"
set "OSGEO4W_FULL=%OSGEO4W_ROOT%"
call "%OSGEO4W_ROOT%\bin\o4w_env.bat"
path "%OSGEO4W_ROOT%\apps\qgis-ltr\bin";"%OSGEO4W_ROOT%\apps\Python312\DLLs";%PATH%
set "QGIS_PREFIX_PATH=%OSGEO4W_ROOT:\=/%/apps/qgis-ltr"
set "QT_QPA_PLATFORM=offscreen"
set "PYTHONPATH=%OSGEO4W_FULL%\apps\qgis-ltr\python"
set "QT_PLUGIN_PATH=%OSGEO4W_ROOT%\apps\qgis-ltr\qtplugins;%OSGEO4W_ROOT%\apps\qt5\plugins;%QT_PLUGIN_PATH%"
echo PYTHONPATH=%PYTHONPATH%
pushd "%OSGEO4W_ROOT%\apps\qgis-ltr\bin"
"%OSGEO4W_ROOT%\bin\python.exe" -c "import qgis.core; print('PyQGIS OK')"
popd
echo Code: %ERRORLEVEL%
pause
