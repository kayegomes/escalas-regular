# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = [
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    "win32timezone",
]
hiddenimports += collect_submodules("openpyxl")

a = Analysis(
    ["gerador_escalas_desktop.py"],
    pathex=["."],
    binaries=[],
    datas=[("app_config.json", ".")],
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name="GeradorEscalas",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="GeradorEscalas",
)
