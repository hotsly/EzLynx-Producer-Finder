# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

a = Analysis(
    ['test.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('chromedriver-win64', 'chromedriver-win64'), # Include the chromedriver directory
        ('icon.ico', 'icon.ico'),  # Include the User Data directory
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Producer Finder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to False to avoid the console window for GUI applications
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'
)
