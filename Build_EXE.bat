@echo off
setlocal
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

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean "SuccessFactorsWeeklyTrainingReport.spec"
if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    pause
    exit /b 1
)

set "APPDIR=dist\Success Factors Weekly Training Report"
if not exist "%APPDIR%\Output" mkdir "%APPDIR%\Output"
if not exist "%APPDIR%\Archive" mkdir "%APPDIR%\Archive"
if not exist "%APPDIR%\Logs" mkdir "%APPDIR%\Logs"
copy /y "VERSION.txt" "%APPDIR%\VERSION.txt" >nul

echo.
echo BUILD COMPLETE.
echo Folder: %CD%\%APPDIR%
echo.
explorer "%CD%\%APPDIR%"
pause
endlocal
