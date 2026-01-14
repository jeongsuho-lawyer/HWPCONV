# -*- mode: python ; coding: utf-8 -*-

# EXE 크기 최소화를 위한 제외 목록
EXCLUDES = [
    # 테스트/개발 모듈
    'pytest', 'unittest', 'doctest', 'pdb', 'pydoc',
    # 사용하지 않는 대형 패키지
    'numpy', 'pandas', 'scipy', 'matplotlib',
    'IPython', 'jupyter', 'notebook',
    # 기타 불필요 모듈
    'lib2to3',
    # tkinter 관련 불필요 항목
    'tkinter.test', 'tkinter.tix',
    # SSL/네트워크 테스트
    'test', 'tests',

    # === google.generativeai SDK 제거 (REST API로 대체) ===
    'google', 'google.generativeai', 'google.ai', 'google.api_core',
    'google.auth', 'google.oauth2', 'google.protobuf',
    'googleapiclient', 'grpc', 'grpcio', 'grpcio_status',
    'proto', 'protobuf', 'httplib2', 'pyparsing',
    'pydantic', 'pydantic_core',

    # === 기타 불필요 패키지 ===
    'tqdm', 'pyreadline3', 'email_validator', 'hypothesis', 'rich', 'dotenv',
    'flask', 'jinja2',  # 웹 서버용 (GUI에서 불필요)
    'cryptography',  # requests에서 선택적 사용
]

a = Analysis(
    ['run_gui.py'],
    pathex=['D:\\Projects\\HWPCONV\\src'],
    binaries=[],
    datas=[
        ('C:\\Users\\jeong\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\customtkinter', 'customtkinter'),
        ('C:\\Users\\jeong\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\tkinterDnD', 'tkinterDnD'),
        ('C:\\Users\\jeong\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\certifi\\cacert.pem', 'certifi'),
    ],
    hiddenimports=[
        'PIL._tkinter_finder',  # PIL 이미지 변환용
        'tkinterDnD',
        'customtkinter',
        'requests',
        'urllib3',
        'certifi',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=2,  # 최대 최적화 (docstrings 제거, assert 제거)
)

# 추가 필터링: 불필요한 바이너리 제거 (SSL은 유지)
a.binaries = [
    (name, path, typ) for name, path, typ in a.binaries
    if not any(x in name.lower() for x in ['qt5', 'qt6', '_test'])
]

pyz = PYZ(a.pure)

# 부트로더 레벨 스플래시 (Python 실행 전에 표시)
splash = Splash(
    'splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,  # 로딩 텍스트 비활성화
    minify_script=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.datas,
    [],
    name='HWP2MD',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=False,  # UPX 끄기 (압축 해제 시간 절약)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
