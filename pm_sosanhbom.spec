# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for SSBOM Manager

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None
project_root = Path.cwd()

hidden_imports = [
    "PyQt6",
    "PyQt6.QtCore",
    "PyQt6.QtGui",
    "PyQt6.QtWidgets",
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    "win32com",
    "win32com.client",
    "pandas",
    "selenium",
    "selenium.webdriver",
    "selenium.webdriver.edge",
    "selenium.webdriver.chrome",
    "psutil",
    "src",
    "src.core",
    "src.automation",
    "src.automation.sap",
    "src.automation.tc14",
    "src.ui",
    "src.gui",
    "src.services",
    "src.reporting",
]

datas = [
    ("locales", "locales"),
    ("update_sources.default.json", "."),
]

a = Analysis(
    ["src/gui/app.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "scipy", "notebooklm"],
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
    name="SSBOM_Portable",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="SSBOM_Portable",
)

