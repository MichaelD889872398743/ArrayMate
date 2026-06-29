"""Build and package the ArrayMate Windows executable."""

from __future__ import annotations

import ast
import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_NAME = "ArrayMate"
SPEC_FILE = ROOT / "ArrayMate.spec"
DIST_APP_DIR = ROOT / "dist" / APP_NAME
DIST_EXE = DIST_APP_DIR / f"{APP_NAME}.exe"
RELEASE_DIR = ROOT / "release"
RELEASE_APP_DIR = RELEASE_DIR / APP_NAME
PORTABLE_PYTHON_DIR = RELEASE_APP_DIR / "python"
PORTABLE_APP_DIR = RELEASE_APP_DIR / "app"
PORTABLE_SITE_PACKAGES_DIR = PORTABLE_APP_DIR / "site-packages"
PYTHON_RUNTIME_FILES = [
    "python.exe",
    "pythonw.exe",
    "python3.dll",
    f"python{sys.version_info.major}{sys.version_info.minor}.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
    "LICENSE.txt",
]
PYTHON_RUNTIME_DIRS = [
    "DLLs",
    "Lib",
]
PYTHON_LIB_EXCLUDES = {
    "__pycache__",
    "site-packages",
    "test",
    "tests",
    "tkinter",
    "turtledemo",
    "idlelib",
    "ensurepip",
    "venv",
    "distutils",
}
THIRD_PARTY_PACKAGES = [
    "shiboken6",
    "openpyxl",
    "et_xmlfile",
]
PYSIDE6_FILES = [
    "__init__.py",
    "__feature__.pyi",
    "_config.py",
    "_git_pyside_version.py",
    "py.typed",
    "pyside6.abi3.dll",
    "pyside6qml.abi3.dll",
    "QtCore.pyd",
    "QtGui.pyd",
    "QtWidgets.pyd",
    "QtNetwork.pyd",
    "Qt6Core.dll",
    "Qt6Gui.dll",
    "Qt6Widgets.dll",
    "Qt6Network.dll",
    "Qt6OpenGL.dll",
    "Qt6OpenGLWidgets.dll",
    "Qt6Svg.dll",
    "Qt6Xml.dll",
    "concrt140.dll",
    "msvcp140.dll",
    "msvcp140_1.dll",
    "msvcp140_2.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
    "opengl32sw.dll",
]
PYSIDE6_PLUGIN_FILES = {
    "platforms": ["qwindows.dll"],
    "styles": ["qwindowsvistastyle.dll"],
    "imageformats": ["qgif.dll", "qico.dll", "qjpeg.dll", "qsvg.dll"],
    "iconengines": ["qsvgicon.dll"],
}


def remove_path(path: Path) -> None:
    """Remove a generated file or directory inside the project root."""
    resolved = path.resolve()
    if ROOT not in (resolved, *resolved.parents):
        raise RuntimeError(f"Refusing to remove path outside project root: {resolved}")
    if path.is_dir():
        shutil.rmtree(path)
        print(f"Removed {path.relative_to(ROOT)}")
    elif path.exists():
        path.unlink()
        print(f"Removed {path.relative_to(ROOT)}")


def clean_generated_outputs() -> None:
    """Remove old build outputs before creating a fresh package."""
    print("Cleaning previous build outputs...")
    for path in [ROOT / "build", ROOT / "dist", RELEASE_DIR]:
        remove_path(path)
    for zip_path in ROOT.glob(f"{APP_NAME}-v*-Windows*.zip"):
        remove_path(zip_path)


def copytree_filtered(source: Path, destination: Path, excluded_names: set[str]) -> None:
    """Copy a directory while skipping well-known bulky/unneeded folders."""
    def ignore(_path: str, names: list[str]) -> set[str]:
        return {name for name in names if name in excluded_names or name.endswith(".pyc")}

    shutil.copytree(source, destination, ignore=ignore)


def read_version() -> str:
    """Read the package version without importing runtime dependencies."""
    init_file = ROOT / "arraymate" / "__init__.py"
    module = ast.parse(init_file.read_text(encoding="utf-8"))
    for node in module.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "__version__":
                    return ast.literal_eval(node.value)
    raise RuntimeError("Could not find __version__ in arraymate/__init__.py")


def validate_build_inputs() -> None:
    required_files = [
        ROOT / "app.py",
        ROOT / "version.txt",
        ROOT / "icon.ico",
        ROOT / "assets" / "arraymate_icon.png",
        ROOT / "assets" / "arraymate_tray_icon.png",
        SPEC_FILE,
    ]
    missing = [path.relative_to(ROOT) for path in required_files if not path.exists()]
    if missing:
        joined = ", ".join(str(path) for path in missing)
        raise FileNotFoundError(f"Missing required build file(s): {joined}")


def ensure_pyinstaller_available() -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "PyInstaller is not installed in this Python environment. "
            "Install build dependencies first, then rerun this script."
        ) from exc


def validate_portable_python_inputs() -> None:
    required_files = [
        ROOT / "app.py",
        ROOT / "arraymate" / "qt_desktop.py",
        ROOT / "requirements.txt",
        ROOT / "assets" / "arraymate_icon.png",
        ROOT / "assets" / "arraymate_tray_icon.png",
    ]
    missing = [path.relative_to(ROOT) for path in required_files if not path.exists()]
    if missing:
        joined = ", ".join(str(path) for path in missing)
        raise FileNotFoundError(f"Missing required portable package file(s): {joined}")


def copy_python_runtime() -> None:
    source_root = Path(sys.prefix)
    print(f"Copying Python runtime from {source_root}...")
    PORTABLE_PYTHON_DIR.mkdir(parents=True)

    for filename in PYTHON_RUNTIME_FILES:
        source = source_root / filename
        if source.exists():
            shutil.copy2(source, PORTABLE_PYTHON_DIR / filename)

    for dirname in PYTHON_RUNTIME_DIRS:
        source = source_root / dirname
        if not source.exists():
            continue
        destination = PORTABLE_PYTHON_DIR / dirname
        if dirname == "Lib":
            copytree_filtered(source, destination, PYTHON_LIB_EXCLUDES)
        else:
            shutil.copytree(source, destination)


def package_location(package_name: str) -> Path:
    import importlib.util

    spec = importlib.util.find_spec(package_name)
    if spec is None or spec.origin is None:
        raise RuntimeError(f"Required package is not installed: {package_name}")
    if spec.submodule_search_locations:
        return Path(next(iter(spec.submodule_search_locations)))
    return Path(spec.origin)


def copy_third_party_packages() -> None:
    print("Copying third-party packages...")
    PORTABLE_SITE_PACKAGES_DIR.mkdir(parents=True)
    copy_pyside6_runtime()
    for package_name in THIRD_PARTY_PACKAGES:
        source = package_location(package_name)
        destination = PORTABLE_SITE_PACKAGES_DIR / source.name
        if source.is_dir():
            shutil.copytree(source, destination, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
        else:
            shutil.copy2(source, destination)
        print(f"Copied {package_name}")


def copy_pyside6_runtime() -> None:
    source = package_location("PySide6")
    destination = PORTABLE_SITE_PACKAGES_DIR / "PySide6"
    destination.mkdir(parents=True)

    for filename in PYSIDE6_FILES:
        source_file = source / filename
        if source_file.exists():
            shutil.copy2(source_file, destination / filename)

    plugins_destination = destination / "plugins"
    for plugin_dir, plugin_files in PYSIDE6_PLUGIN_FILES.items():
        source_dir = source / "plugins" / plugin_dir
        destination_dir = plugins_destination / plugin_dir
        destination_dir.mkdir(parents=True, exist_ok=True)
        for filename in plugin_files:
            source_file = source_dir / filename
            if source_file.exists():
                shutil.copy2(source_file, destination_dir / filename)

    print("Copied PySide6 runtime subset")


def copy_application_files() -> None:
    print("Copying ArrayMate application files...")
    PORTABLE_APP_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ROOT / "app.py", PORTABLE_APP_DIR / "app.py")
    shutil.copytree(ROOT / "arraymate", PORTABLE_APP_DIR / "arraymate", ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    shutil.copytree(ROOT / "assets", PORTABLE_APP_DIR / "assets")
    for source in [
        ROOT / "README.md",
        ROOT / "LICENSE",
        ROOT / "CHANGELOG.md",
        *sorted(ROOT.glob("sample_data*.json")),
    ]:
        if source.exists():
            shutil.copy2(source, RELEASE_APP_DIR / source.name)


def write_portable_launchers() -> None:
    launcher = RELEASE_APP_DIR / "Run ArrayMate.bat"
    launcher.write_text(
        "\n".join(
            [
                "@echo off",
                "setlocal",
                'set "APP_DIR=%~dp0app"',
                'set "PYTHONHOME=%~dp0python"',
                'set "PYTHONPATH=%APP_DIR%;%APP_DIR%\\site-packages"',
                'start "" "%~dp0python\\pythonw.exe" "%APP_DIR%\\app.py"',
                "",
            ]
        ),
        encoding="utf-8",
    )

    notes = RELEASE_APP_DIR / "PORTABLE_README.txt"
    notes.write_text(
        "\n".join(
            [
                "ArrayMate Portable Python Package",
                "=================================",
                "",
                "Run ArrayMate.bat starts the app through the bundled Python runtime.",
                "",
                "This package intentionally uses pythonw.exe instead of a generated",
                "ArrayMate.exe. Some managed Windows laptops block unsigned generated",
                "executables, while still allowing trusted Python runtimes to start.",
                "",
                "Windows may still show an Unknown publisher or SmartScreen warning",
                "because ArrayMate is not code-signed. This is expected for the",
                "current open-source release.",
                "",
                "Keep the python and app folders next to Run ArrayMate.bat.",
                "",
            ]
        ),
        encoding="utf-8",
    )


def create_zip_from_release(version: str, suffix: str) -> Path:
    zip_path = ROOT / f"{APP_NAME}-v{version}-{suffix}.zip"
    print(f"Creating {zip_path.name}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in sorted(RELEASE_DIR.rglob("*")):
            if file_path.is_file():
                zip_file.write(file_path, file_path.relative_to(RELEASE_DIR))
    print(f"Release package ready: {zip_path.name}")
    return zip_path


def build_executable() -> None:
    print("Building portable ArrayMate folder from ArrayMate.spec...")
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--clean", "--noconfirm", str(SPEC_FILE)],
        cwd=ROOT,
        check=True,
    )
    if not DIST_EXE.exists():
        raise FileNotFoundError(f"Expected executable was not created: {DIST_EXE}")


def create_release_package(version: str) -> None:
    print("Creating release folder...")
    RELEASE_DIR.mkdir()

    shutil.copytree(DIST_APP_DIR, RELEASE_APP_DIR)
    print(f"Copied {RELEASE_APP_DIR.relative_to(ROOT)}")

    for source in [
        ROOT / "README.md",
        ROOT / "LICENSE",
        ROOT / "CHANGELOG.md",
        *sorted(ROOT.glob("sample_data*.json")),
    ]:
        if source.exists():
            destination = RELEASE_APP_DIR / source.name
            shutil.copy2(source, destination)
            print(f"Copied {destination.relative_to(ROOT)}")

    create_zip_from_release(version, "Windows-PyInstaller")


def create_portable_python_package(version: str) -> None:
    print("Creating portable Python package...")
    RELEASE_DIR.mkdir()
    RELEASE_APP_DIR.mkdir()
    copy_python_runtime()
    copy_third_party_packages()
    copy_application_files()
    write_portable_launchers()
    create_zip_from_release(version, "Windows-PortablePython")


def cleanup() -> None:
    """Remove generated build folders while keeping source-controlled spec files."""
    for path in [ROOT / "build", ROOT / "__pycache__"]:
        remove_path(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build ArrayMate release packages.")
    parser.add_argument(
        "target",
        nargs="?",
        choices=("portable-python", "pyinstaller"),
        default="portable-python",
        help="Package target to build. Default: portable-python.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    version = read_version()
    print(f"Starting {APP_NAME} v{version} {args.target} build")
    print("=" * 50)

    try:
        clean_generated_outputs()
        if args.target == "pyinstaller":
            validate_build_inputs()
            ensure_pyinstaller_available()
            build_executable()
            create_release_package(version)
        else:
            validate_portable_python_inputs()
            create_portable_python_package(version)
    except Exception as exc:
        print(f"Build failed: {exc}")
        return 1

    print("=" * 50)
    print("Build completed successfully.")
    if args.target == "pyinstaller":
        print(f"Executable: {RELEASE_APP_DIR / f'{APP_NAME}.exe'}")
    else:
        print(f"Launcher: {RELEASE_APP_DIR / 'Run ArrayMate.bat'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
