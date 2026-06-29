@echo off
setlocal

echo Building ArrayMate portable Python package...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH.
    pause
    exit /b 1
)

python build_exe.py portable-python
set BUILD_EXIT_CODE=%ERRORLEVEL%

echo.
if %BUILD_EXIT_CODE% EQU 0 (
    echo Build process completed.
    echo Run release\ArrayMate\Run ArrayMate.bat.
) else (
    echo Build process failed with exit code %BUILD_EXIT_CODE%.
)

pause
exit /b %BUILD_EXIT_CODE%
