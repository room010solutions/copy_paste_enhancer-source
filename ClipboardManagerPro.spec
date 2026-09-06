# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# customtkinter ships theme/asset JSON files that PyInstaller won't find
# automatically — collect_data_files pulls them in explicitly.
datas = collect_data_files("customtkinter")

a = Analysis(
    ['copy_paste_02.py'],       # rename here if you rename your main script
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['win32clipboard', 'win32con'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ClipboardManagerPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,              # windowed app — no console popup
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',        # remove this line if you don't have an icon
)