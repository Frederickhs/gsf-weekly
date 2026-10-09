@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo ERROR: The project virtual environment was not found.
    echo Expected: %CD%\.venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

if not exist "Scripts\gsf_app.py" (
    echo.
    echo ERROR: Scripts\gsf_app.py was not found.
    echo.
    pause
    exit /b 1
)

start "Success Factors Weekly Training Report" ".venv\Scripts\pythonw.exe" "Scripts\gsf_app.py"
endlocal
