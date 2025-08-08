"""
Build script for ArrayMate executable
Automates the creation of a standalone .exe file
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_pyinstaller():
    """Install PyInstaller if not already installed."""
    try:
        import PyInstaller
        print("PyInstaller is already installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("PyInstaller installed successfully")

def build_executable():
    """Build the executable using PyInstaller."""
    print("Building ArrayMate executable...")
    
    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",                    
        "--windowed",                   
        "--name=ArrayMate",             
        "--icon=icon.ico",
        "--version-file=version.txt",              
        "--add-data=sample_data.json;.", 
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "array_mate.py"
    ]
    
    # Remove icon if it doesn't exist
    if not os.path.exists("icon.ico"):
        cmd.remove("--icon=icon.ico")
    
    # Run PyInstaller
    try:
        subprocess.run(cmd, check=True)
        print("Executable built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error building executable: {e}")
        return False

def create_release_package():
    """Create a release package with the executable and necessary files."""
    print("Creating release package...")
    
    # Create release directory
    release_dir = Path("release")
    release_dir.mkdir(exist_ok=True)
    
    # Copy executable
    exe_path = Path("dist/ArrayMate.exe")
    if exe_path.exists():
        shutil.copy2(exe_path, release_dir / "ArrayMate.exe")
        print("Copied executable to release folder")
    else:
        print("Executable not found!")
        return False
    
    # Copy sample data
    sample_data = Path("sample_data.json")
    if sample_data.exists():
        shutil.copy2(sample_data, release_dir / "sample_data.json")
        print("Copied sample data to release folder")
    
    # Copy README
    readme = Path("README.md")
    if readme.exists():
        shutil.copy2(readme, release_dir / "README.md")
        print("Copied README to release folder")
    
    # Copy LICENSE
    license_file = Path("LICENSE")
    if license_file.exists():
        shutil.copy2(license_file, release_dir / "LICENSE")
        print("Copied LICENSE to release folder")
    
    # Create requirements.txt for reference
    requirements = Path("requirements.txt")
    if requirements.exists():
        shutil.copy2(requirements, release_dir / "requirements.txt")
        print("Copied requirements.txt to release folder")
    
    print("Release package created successfully!")
    return True

def create_zip_package():
    """Create a zip file for GitHub release."""
    import zipfile
    
    print("Creating zip package...")
    
    release_dir = Path("release")
    zip_name = "ArrayMate-v1.0.0-Windows.zip"
    
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in release_dir.rglob("*"):
            if file_path.is_file():
                arcname = file_path.relative_to(release_dir)
                zipf.write(file_path, arcname)
                print(f"Added {arcname} to zip")
    
    print(f"Zip package created: {zip_name}")
    return zip_name

def cleanup():
    """Clean up build artifacts."""
    print("Cleaning up build artifacts...")
    
    # Remove PyInstaller build directories
    build_dirs = ["build", "dist", "__pycache__"]
    for dir_name in build_dirs:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"Removed {dir_name}")
    
    # Remove .spec file
    spec_file = "ArrayMate.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print("Removed .spec file")

def main():
    """Main build process."""
    print("Starting ArrayMate build process...")
    print("=" * 50)
    
    # Step 1: Install PyInstaller
    install_pyinstaller()
    
    # Step 2: Build executable
    if not build_executable():
        print("Build failed!")
        return False
    
    # Step 3: Create release package
    if not create_release_package():
        print("Release package creation failed!")
        return False
    
    # Step 4: Create zip package
    zip_name = create_zip_package()
    
    # Step 5: Cleanup (optional - comment out if you want to keep build files)
    # cleanup()
    
    print("=" * 50)
    print("Build completed successfully!")
    print(f"Release package: {zip_name}")
    print("Executable location: release/ArrayMate.exe")
    print("\n Next steps:")
    print("1. Test the executable: release/ArrayMate.exe")
    print("2. Upload the zip file to GitHub releases")
    print("3. Tag the release with v1.0.0")
    
    return True

if __name__ == "__main__":
    main()
