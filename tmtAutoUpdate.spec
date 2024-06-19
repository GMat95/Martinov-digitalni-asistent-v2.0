# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('template.xlsx', 'new'), ('timtecLogo.ico', 'new'), ('tmtblack.png', 'new')],
    hiddenimports=['uuid', 'pymssql', 'pillow'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
a.datas += [('timtecLogo.ico', 'C:\\Users\\TMT-IT-NTB-05\\Desktop\\TimtecAutoFolder\\new\\timtecLogo.ico', 'DATA')]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Timtec folder automation - Martinov digitalni asistent',
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
    icon=['C:\\Users\\TMT-IT-NTB-05\\Desktop\\TimtecAutoFolder\\new\\timtecLogo.ico'],
)
