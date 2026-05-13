# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path

tkdnd_src = Path(r'C:\Users\user\AppData\Local\Programs\Python\Python313\Lib\site-packages\tkinterdnd2\tkdnd')

a = Analysis(
    ['compare_bom_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        (str(tkdnd_src), 'tkinterdnd2/tkdnd'),
    ],
    hiddenimports=[
        'tkinterdnd2',
        'openpyxl',
        'pandas',
    ],
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
    name='BOM_Comparer',
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
