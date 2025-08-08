@echo off
echo Building ArrayMate executable...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Run the build script
python build_exe.py

echo.
echo Build process completed!
echo Check the 'release' folder for the executable.
pause
