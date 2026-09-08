# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['fitz', 'PIL', 'openpyxl']
hiddenimports += collect_submodules('core')


a = Analysis(
    ['C:/Users/Ruth/Desktop/价格牌处理/kaifa/main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/Ruth/Desktop/价格牌处理/kaifa/samples', 'samples'), ('C:/Users/Ruth/Desktop/价格牌处理/kaifa/assets', 'assets')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy'],
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
    name='价格牌处理工具',
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
    icon=['C:/Users/Ruth/Desktop/价格牌处理/kaifa/assets/logo.ico'],
)
