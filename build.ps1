# ArrayMate build launcher for Windows PowerShell.

$ErrorActionPreference = "Stop"

Write-Host "ArrayMate build" -ForegroundColor Green
Write-Host "===============" -ForegroundColor Green

try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: Python is not installed or not in PATH." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$requiredFiles = @(
    "app.py",
    "ArrayMate.spec",
    "version.txt",
    "icon.ico",
    "assets\arraymate_icon.png",
    "assets\arraymate_tray_icon.png"
)

foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        Write-Host "Missing required file: $file" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

python build_exe.py portable-python
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Host ""
    Write-Host "Build completed successfully." -ForegroundColor Green
    if (Test-Path "release\ArrayMate\Run ArrayMate.bat") {
        Write-Host "Launcher: release\ArrayMate\Run ArrayMate.bat" -ForegroundColor Green
    }
    $zipFiles = Get-ChildItem "ArrayMate-v*-Windows-*.zip" -ErrorAction SilentlyContinue
    if ($zipFiles) {
        Write-Host "Release package: $($zipFiles[0].Name)" -ForegroundColor Green
    }
} else {
    Write-Host ""
    Write-Host "Build failed with exit code $exitCode." -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"
exit $exitCode
