# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['run_gui.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        ('src/hwpconv', 'hwpconv'),
        ('C:/Users/jeong/AppData/Local/Programs/Python/Python312/Lib/site-packages/tkinterDnD/windows', 'tkdnd2.9.2'),
    ],
    hiddenimports=['hwpconv', 'hwpconv.config', 'hwpconv.image_analyzer', 'hwpconv.parsers', 'hwpconv.parsers.hwp', 'hwpconv.parsers.hwpx', 'hwpconv.converters', 'hwpconv.converters.base', 'hwpconv.converters.html', 'hwpconv.converters.markdown', 'hwpconv.models', 'hwpconv.utils'],
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
    name='HWP2MD',
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
