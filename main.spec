# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import copy_metadata

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
ICON_FILE = os.path.join(BASE_DIR, "icon.ico")

added_datas = []
added_datas += copy_metadata("imageio")
added_datas += copy_metadata("imageio_ffmpeg")
added_datas += copy_metadata("moviepy")
added_datas += copy_metadata("pydub")

a = Analysis(
    ['main.py'],
    pathex=[BASE_DIR],
    binaries=[],
    datas=added_datas,
    hiddenimports=[
        'imageio',
        'imageio.v2',
        'imageio.plugins',
        'imageio_ffmpeg',
        'moviepy',
        'moviepy.video',
        'moviepy.audio',
        'pydub',
    ],
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
    [],
    exclude_binaries=True,
    name='SilenceCut',
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
    icon=ICON_FILE,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SilenceCut',
)