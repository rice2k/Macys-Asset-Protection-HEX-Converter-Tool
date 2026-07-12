# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

assets_dir = Path('src') / 'assets'
datas = []
if assets_dir.exists():
    datas.append((str(assets_dir), 'assets'))

binaries = []
hiddenimports = []
try:
    collected = collect_all('tkinterdnd2')
    datas += collected[0]
    binaries += collected[1]
    hiddenimports += collected[2]
except Exception:
    pass

icon_path = assets_dir / 'app-icon.ico'
icon_value = [str(icon_path)] if icon_path.exists() else None

a = Analysis(
    ['desktop_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='AP_HEX_Converter_Tool_v1.0.3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_value,
)
