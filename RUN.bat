@echo off
title YASKAWA Import Accounting & Tally Automation
cd /d "%~dp0"

set "PYTHON_CMD=py312\python.exe"
if not exist "%PYTHON_CMD%" (
    set "PYTHON_CMD=python"
)

if not "%~1"=="" (
    "%PYTHON_CMD%" main.py %*
    goto :end
)

:menu
cls
echo ===============================================================================
echo        YASKAWA IMPORT ACCOUNTING & TALLY EXCEL AUTOMATION APPLICATION
echo ===============================================================================
echo.
echo   [1] Launch Interactive Interface (Drop files, inspect status, view checks)
echo   [2] Run Direct Pipeline (Automatic execution in console)
echo   [3] Run Automated Test Suite (pytest verification)
echo   [4] Open Input Folder (input\)
echo   [5] Open Output Folder (output\)
echo   [6] Exit
echo.
echo ===============================================================================
set /p choice="Select an option [1-6] (Default is 1): "

if "%choice%"=="" set choice=1
if "%choice%"=="1" goto :launch_gui
if "%choice%"=="2" goto :run_cli
if "%choice%"=="3" goto :run_test
if "%choice%"=="4" goto :open_input
if "%choice%"=="5" goto :open_output
if "%choice%"=="6" goto :exit

:launch_gui
echo.
echo Launching Interactive Desktop Interface...
"%PYTHON_CMD%" main.py --web
goto :menu

:run_cli
echo.
echo Running Accounting Pipeline directly...
"%PYTHON_CMD%" main.py --no-gui
echo.
pause
goto :menu

:run_test
echo.
echo Running automated verification test suite...
"%PYTHON_CMD%" -m pytest test_automation.py -v
echo.
pause
goto :menu

:open_input
start "" "input"
goto :menu

:open_output
start "" "output"
goto :menu

:exit
exit /b 0

:end
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application ended with error code %ERRORLEVEL%.
    pause
)
