@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo Building Success Factors Weekly Training Report...
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: .venv\Scripts\python.exe was not found.
    pause
    exit /b 1
)

if not exist "SuccessFactorsWeeklyTrainingReport.spec" (
    echo ERROR: SuccessFactorsWeeklyTrainingReport.spec was not found.
    pause
    exit /b 1
)

set "APPNAME=Success Factors Weekly Training Report"
set "APPDIR=dist\%APPNAME%"
set "BUILDARCHIVE=Archive\Application Builds"

if exist "%APPDIR%" (
    for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmmss"') do set "TIMESTAMP=%%I"
    if not exist "%BUILDARCHIVE%" mkdir "%BUILDARCHIVE%"

    echo Archiving previous application build...
    move "%APPDIR%" "%BUILDARCHIVE%\%APPNAME%_!TIMESTAMP!" >nul

    if errorlevel 1 (
        echo.
        echo ERROR: The previous application folder could not be archived.
        echo Close the EXE and any files opened from the distribution folder, then try again.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean "SuccessFactorsWeeklyTrainingReport.spec"
if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    pause
    exit /b 1
)

if not exist "%APPDIR%\Output" mkdir "%APPDIR%\Output"
if not exist "%APPDIR%\Archive" mkdir "%APPDIR%\Archive"
if not exist "%APPDIR%\Logs" mkdir "%APPDIR%\Logs"
copy /y "VERSION.txt" "%APPDIR%\VERSION.txt" >nul
copy /y "README.txt" "%APPDIR%\README.txt" >nul

echo.
echo BUILD COMPLETE.
echo Folder: %CD%\%APPDIR%
echo Previous builds: %CD%\%BUILDARCHIVE%
echo.
explorer "%CD%\%APPDIR%"
pause
endlocal
