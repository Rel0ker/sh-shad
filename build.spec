# -*- mode: python ; coding: utf-8 -*-
"""Один файл: pyinstaller build.spec → dist/schedule_changes.exe"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_all

block_cipher = None
root = Path(SPECPATH).resolve()

pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

a = Analysis(
    [str(root / "app.py")],
    pathex=[str(root)],
    binaries=pw_binaries,
    datas=[
        (str(root / "templates"), "templates"),
        (str(root / "static"), "static"),
        (str(root / "data"), "data"),
    ]
    + pw_datas,
    hiddenimports=list(pw_hiddenimports)
    + [
        "playwright.sync_api",
        "playwright._impl._driver",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(root / "rthook_playwright.py")],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="schedule_changes",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
