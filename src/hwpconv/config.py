"""
설정 관리 모듈
- API 키 등 사용자 설정을 로컬에 안전하게 저장/로드
"""
import os
import json
from pathlib import Path


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
