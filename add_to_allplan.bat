@echo off
REM This script will place symbolic link in your ALLLPLAN Usr or Std directory pointing to this repository
REM for testing purposes
set /p action="Do you want to install (i) or remove (R) the links? "
set /p targetPath="Please enter the Path to Usr or Std (with trailing \): "


if exist "%targetPath%Library\Allplan GmbH\ExcelReport.pyp" (
    del "%targetPath%Library\Allplan GmbH\ExcelReport.pyp"
)

if exist "%targetPath%Library\Allplan GmbH\ExcelReport.svg" (
    del "%targetPath%Library\Allplan GmbH\ExcelReport.svg"
)

if exist "%targetPath%PythonPartsScripts\allplan_gmbh\ExcelReport.py" (
    del "%targetPath%PythonPartsScripts\allplan_gmbh\ExcelReport.py"
)

echo "Removal process completed"


if /I "%action%"=="i" (
    goto :install
) else if /I "%action%"=="R" (
    echo "Press any key to continue."
    pause >null
    exit /b 0
)

:install

set scriptDir=%~dp0

if not exist "%targetPath%\PythonPartsScripts" (
    mkdir "%targetPath%\PythonPartsScripts"
)

if not exist "%targetPath%\Library\Allplan GmbH" (
    mkdir "%targetPath%\Library\Allplan GmbH"
)

if not exist "%targetPath%\PythonPartsScripts\allplan_gmbh" (
    mkdir "%targetPath%\PythonPartsScripts\allplan_gmbh"
)

mklink "%targetPath%Library\Allplan GmbH\ExcelReport.svg" "%scriptDir%Library\Allplan GmbH\ExcelReport.svg"
mklink "%targetPath%Library\Allplan GmbH\ExcelReport.pyp" "%scriptDir%Library\Allplan GmbH\ExcelReport.pyp"
mklink "%targetPath%PythonPartsScripts\allplan_gmbh\ExcelReport.py" "%scriptDir%PythonPartsScripts\allplan_gmbh\ExcelReport.py"

echo "PythonPart installed in Allplan. You'll find it in Library -> Office or Private -> Plugin Hub."
echo "Press any key to continue"
pause >nul
exit /b 0