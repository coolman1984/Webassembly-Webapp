@echo off
rem Starts the dashboard. Nothing needs to be installed: it uses the "runtime" folder inside this project.
rem pushd (instead of cd) also works when the folder is on a network share.
setlocal
pushd "%~dp0"
title BOM Confirmation Plan dashboard

if not exist "runtime\python.exe" (
    echo.
    echo  The "runtime" folder is missing from this project folder.
    echo  Please copy the WHOLE project folder, not only some of its files.
    echo.
    pause
    popd
    exit /b 1
)

"runtime\python.exe" -I -X utf8 server.py
if errorlevel 1 (
    echo.
    echo  The dashboard stopped unexpectedly. Details are in: data\pipeline.log
    echo.
    pause
)
popd
