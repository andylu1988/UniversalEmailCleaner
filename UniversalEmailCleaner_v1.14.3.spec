# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['universal_email_cleaner.py'],
    pathex=[],
    binaries=[],
    datas=[('graph-mail-delete.ico', '.'), ('avatar_b64.txt', '.'), ('license_manager.py', '.')],
    hiddenimports=['license_manager'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='UniversalEmailCleaner_v1.14.3',
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
    icon=['graph-mail-delete.ico'],
)
