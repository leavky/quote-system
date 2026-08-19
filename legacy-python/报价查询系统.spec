# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


ROOT = Path.cwd()


a = Analysis(
    ["ui.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=["pandas", "openpyxl", "webview"],
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
    exclude_binaries=False,
    name="quote-system",
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
