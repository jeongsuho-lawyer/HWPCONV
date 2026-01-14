"""
설정 관리 모듈
- API 키 등 사용자 설정을 로컬에 안전하게 저장/로드
"""
import os
import json
from pathlib import Path

# 기본 제공 Gemini 모델 목록
# (model_id, 표시명, 설명)
# 무료 티어 제한은 변동될 수 있음 - 공식 문서 참조: https://ai.google.dev/gemini-api/docs/rate-limits
GEMINI_MODELS = [
    ("gemini-3-flash-preview", "Gemini 3 Flash", "최신 고성능"),
    ("gemini-2.5-flash", "Gemini 2.5 Flash", "균형잡힌 성능"),
    ("gemini-2.5-flash-lite", "Gemini 2.5 Flash Lite", "초고속"),
    ("gemini-2.0-flash", "Gemini 2.0 Flash", "안정적"),
]

# 무료 티어 제한 정보 (분당 이미지 분석 횟수)
GEMINI_FREE_LIMITS = {
    "gemini-3-flash-preview": 10,
    "gemini-2.5-flash": 15,
    "gemini-2.5-flash-lite": 30,
    "gemini-2.0-flash": 15,
}

DEFAULT_MODEL = "gemini-3-flash-preview"


def get_config_dir() -> Path:
    """설정 파일 디렉토리 반환 (Windows: %APPDATA%\HwpConverter)"""
    if os.name == 'nt':  # Windows
        base = os.environ.get('APPDATA', os.path.expanduser('~'))
    else:  # macOS/Linux
        base = os.path.expanduser('~/.config')
    
    config_dir = Path(base) / 'HwpConverter'
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def get_config_path() -> Path:
    """설정 파일 경로 반환"""
    return get_config_dir() / 'config.json'


def load_config() -> dict:
    """설정 파일 로드"""
    config_path = get_config_path()
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(config: dict):
    """설정 파일 저장"""
    config_path = get_config_path()
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def get_api_key() -> str:
    """저장된 Gemini API 키 반환"""
    config = load_config()
    return config.get('gemini_api_key', '')


def save_api_key(api_key: str):
    """Gemini API 키 저장"""
    config = load_config()
    config['gemini_api_key'] = api_key
    save_config(config)


def has_api_key() -> bool:
    """API 키가 설정되어 있는지 확인"""
    return bool(get_api_key())


def get_model() -> str:
    """저장된 Gemini 모델 ID 반환 (없으면 기본값)"""
    config = load_config()
    return config.get('gemini_model', DEFAULT_MODEL)


def save_model(model_id: str):
    """Gemini 모델 ID 저장"""
    config = load_config()
    config['gemini_model'] = model_id
    save_config(config)


def get_model_display_name(model_id: str) -> str:
    """모델 ID에 해당하는 표시명 반환"""
    for mid, name, _ in GEMINI_MODELS:
        if mid == model_id:
            return name
    # 사용자 정의 모델인 경우 ID 그대로 반환
    return model_id
