# -*- mode: python ; coding: utf-8 -*-
# PyInstaller-Spec-Datei für excel-editor.exe
#
# Verwendung (Windows PowerShell / CMD):
#   pip install pyinstaller
#   pip install -e .
#   pyinstaller excel_editor.spec --clean
#
# Ergebnis: dist/excel-editor.exe  (standalone, keine Python-Installation nötig)

import os

# SPECPATH ist eine PyInstaller-Variable: Verzeichnis der .spec-Datei
src_dir = os.path.join(SPECPATH, "src")
# cli_script = os.path.join(SPECPATH, "src", "excel_editor", "cli.py")
cli_script = os.path.join(SPECPATH, "src", "excel_editor", "__main__.py")

a = Analysis(
    [cli_script],
    pathex=[src_dir],
    binaries=[],
    datas=[],
    hiddenimports=[
        "excel_editor",
        "excel_editor.cli",
        "excel_editor.editor",
        "excel_editor.models",
        # xlwings Windows COM Backend
        "xlwings",
        "xlwings._xlwindows",
        "xlwings.utils",
        "xlwings.constants",
        # pywin32 (COM / Windows API)
        "win32com",
        "win32com.client",
        "win32api",
        "win32con",
        "pythoncom",
        "pywintypes",
        # pydantic
        "pydantic",
        "pydantic_core",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="excel-editor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    # console=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
