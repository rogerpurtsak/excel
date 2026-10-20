# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the Excel AI Solver menu bar app.

Build with:  pyinstaller "Excel AI Solver.spec" --clean
Produces:    dist/Excel AI Solver.app  (menu-bar-only, LSUIElement)
"""

block_cipher = None


a = Analysis(
    ["run_app.py"],
    pathex=["."],
    binaries=[],
    datas=[
        ("resources/icon.png", "resources"),
    ],
    hiddenimports=[
        "rumps",
        "keyring.backends.macOS",
        "anthropic",
        "scipy._lib.array_api_compat.numpy.fft",
        "scipy.special._cdflib",
        # settings_window.py uses these AppKit classes directly
        "AppKit",
        "Foundation",
        "src.settings_window",
        "src.excel_writer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Excel AI Solver",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="Excel AI Solver",
)

app = BUNDLE(
    coll,
    name="Excel AI Solver.app",
    icon=None,
    bundle_identifier="com.excelaisolver.app",
    info_plist={
        "LSUIElement": True,  # menu bar only, no Dock icon
        "CFBundleIdentifier": "com.excelaisolver.app",
        "CFBundleName": "Excel AI Solver",
        "CFBundleDisplayName": "Excel AI Solver",
        "CFBundleShortVersionString": "1.0.0",
        "NSAppleEventsUsageDescription": (
            "Excel AI Solver controls Microsoft Excel to read the active workbook "
            "and write computed statistics results."
        ),
        "NSHighResolutionCapable": True,
    },
)
