# -*- mode: python ; coding: utf-8 -*-


datas = [
    ("assets", "assets"),
]
binaries = []
hiddenimports = [
    "openpyxl",
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
]
excludes = [
    "IPython",
    "PIL.ImageQt",
    "PyQt5",
    "PyQt6",
    "PySide2",
    "astroid",
    "bokeh",
    "boto3",
    "botocore",
    "cryptography",
    "django",
    "docutils",
    "fsspec",
    "jinja2",
    "llvmlite",
    "lxml",
    "matplotlib",
    "numba",
    "numexpr",
    "notebook",
    "odf",
    "pyarrow",
    "pygments",
    "pytest",
    "pythoncom",
    "pywintypes",
    "requests",
    "s3fs",
    "scipy",
    "sphinx",
    "tables",
    "tkinter",
    "urllib3",
    "win32com",
    "xarray",
    "xlrd",
    "xlsxwriter",
]


a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ArrayMate",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="icon.ico",
    version="version.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ArrayMate",
)
