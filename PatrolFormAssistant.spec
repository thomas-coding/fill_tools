# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_main.py'],
    pathex=[],
    binaries=[],
    datas=[('AutoHotkey-v2\\AutoHotkey64.exe', 'AutoHotkey-v2'), ('AutoHotkey-v2\\AutoHotkey32.exe', 'AutoHotkey-v2')],
    hiddenimports=[],
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
    name='PatrolFormAssistant',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
