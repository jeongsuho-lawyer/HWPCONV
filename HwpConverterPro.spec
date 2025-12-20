# -*- mode: python ; coding: utf-8 -*-

# EXE 크기 최소화를 위한 제외 목록
EXCLUDES = [
    # 테스트/개발 모듈
    'pytest', 'unittest', 'doctest', 'pdb', 'pydoc',
    # 사용하지 않는 대형 패키지
    'numpy', 'pandas', 'scipy', 'matplotlib',
    'IPython', 'jupyter', 'notebook',
    # 기타 불필요 모듈
    'xml.etree.ElementTree.doctest',
    'lib2to3', 'distutils', 'setuptools',
    'pkg_resources',
    # tkinter 관련 불필요 항목
    'tkinter.test', 'tkinter.tix',
    # SSL/네트워크 테스트
    'test', 'tests',
]

a = Analysis(
    ['run_gui.py'],
    pathex=['D:\\Projects\\HWPCONV\\src'],
    binaries=[],
    datas=[
        ('C:\\Users\\jeong\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\customtkinter', 'customtkinter'),
        ('C:\\Users\\jeong\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\tkinterDnD', 'tkinterDnD')
    ],
    hiddenimports=[
        'PIL._tkinter_finder',  # PIL 이미지 변환용
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=2,  # 최대 최적화 (docstrings 제거, assert 제거)
)

# 추가 필터링: 불필요한 바이너리 제거
a.binaries = [
    (name, path, typ) for name, path, typ in a.binaries
    if not any(x in name.lower() for x in ['qt5', 'qt6', 'libssl', 'libcrypto', '_test'])
]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='HwpConverterPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,  # 디버그 심볼 제거
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
