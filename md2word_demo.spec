# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_PATH = ROOT / "examples" / "desktop_app" / "app.py"
INDEX_HTML = ROOT / "examples" / "desktop_app" / "index.html"

a = Analysis(
    [str(APP_PATH)],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=[(str(INDEX_HTML), ".")],
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
    name='md2word_demo',
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
