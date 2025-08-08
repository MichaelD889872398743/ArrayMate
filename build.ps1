# ArrayMate Build Script (PowerShell)
# Builds the executable and creates a release package

Write-Host "🚀 ArrayMate Build Script" -ForegroundColor Green
Write-Host "=========================" -ForegroundColor Green

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✓ Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Error: Python is not installed or not in PATH" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if required files exist
$requiredFiles = @("array_mate.py", "requirements.txt", "sample_data.json")
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "✓ Found $file" -ForegroundColor Green
    } else {
        Write-Host "❌ Missing required file: $file" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host "`nStarting build process..." -ForegroundColor Yellow

# Run the build script
try {
    python build_exe.py
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n🎉 Build completed successfully!" -ForegroundColor Green
        
        # Check if executable was created
        if (Test-Path "release\ArrayMate.exe") {
            $fileSize = (Get-Item "release\ArrayMate.exe").Length
            $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
            Write-Host "✓ Executable created: release\ArrayMate.exe ($fileSizeMB MB)" -ForegroundColor Green
        }
        
        # Check if zip was created
        $zipFiles = Get-ChildItem "ArrayMate-v*.zip" -ErrorAction SilentlyContinue
        if ($zipFiles) {
            Write-Host "✓ Release package created: $($zipFiles[0].Name)" -ForegroundColor Green
        }
        
        Write-Host "`n📋 Next steps:" -ForegroundColor Cyan
        Write-Host "1. Test the executable: release\ArrayMate.exe"
        Write-Host "2. Upload the zip file to GitHub releases"
        Write-Host "3. Tag the release with v1.0.0"
        
    } else {
        Write-Host "`n❌ Build failed with exit code $LASTEXITCODE" -ForegroundColor Red
    }
} catch {
    Write-Host "`n❌ Error during build: $_" -ForegroundColor Red
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
