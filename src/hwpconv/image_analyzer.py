"""
Gemini Vision API를 사용한 이미지 분석 모듈 (REST API 직접 호출)
"""
import base64
import json
from typing import Optional
import traceback

from . import config as app_config

# requests는 lazy import (구동 시간 최적화)

# 지원되는 MIME 타입 (Gemini API)
SUPPORTED_MIME_TYPES = {'image/png', 'image/jpeg', 'image/webp', 'image/heic', 'image/heif'}

# Gemini API 엔드포인트
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent"

# 로그 파일 최대 크기 (1MB)
MAX_LOG_SIZE = 1 * 1024 * 1024

def _get_log_path() -> str:
    """로그 파일 경로 반환 (EXE 호환)"""
    log_path = app_config.get_config_dir() / 'gemini_debug.log'

    # 로그 파일 크기 체크 및 rotation
    try:
        if log_path.exists() and log_path.stat().st_size > MAX_LOG_SIZE:
            backup_path = app_config.get_config_dir() / 'gemini_debug.old.log'
            if backup_path.exists():
                backup_path.unlink()
            log_path.rename(backup_path)
    except Exception:
        pass

    return str(log_path)


def is_available() -> bool:
    """이미지 분석 기능 사용 가능 여부"""
    try:
        import requests
        return app_config.has_api_key()
    except ImportError:
        return False


def analyze_image(image_bytes: bytes, mime_type: str = "image/png") -> Optional[str]:
    """
    이미지를 분석하여 설명 텍스트 반환

    Args:
        image_bytes: 이미지 바이너리 데이터
        mime_type: 이미지 MIME 타입 (image/png, image/jpeg 등)

    Returns:
        이미지 설명 문자열 또는 None (실패 시)
    """
    import time

    if not is_available():
        return None

    start_time = time.time()
    original_mime = mime_type

    try:
        api_key = app_config.get_api_key()

        # BMP, TIFF 등 지원되지 않는 포맷은 PNG로 변환
        if mime_type not in SUPPORTED_MIME_TYPES:
            try:
                from PIL import Image as PilImage
                import io

                img = PilImage.open(io.BytesIO(image_bytes))
                output = io.BytesIO()
                img.save(output, format='PNG')
                image_bytes = output.getvalue()
                mime_type = 'image/png'
                print(f"이미지 포맷 변환: {original_mime} → {mime_type}")
            except Exception as e:
                print(f"이미지 변환 실패: {e}")

        print(f"이미지 분석 시작 (크기: {len(image_bytes)} bytes, 타입: {mime_type})...")

        # 이미지를 base64로 인코딩
        image_data = base64.b64encode(image_bytes).decode('utf-8')

        # 프롬프트
        prompt = """이 이미지를 분석해주세요:
- 로고, 아이콘, 심볼 등 단순한 이미지면: 한 줄로 짧게 설명
- 도표, 그래프, 다이어그램, 사진 등 내용이 있는 이미지면: 최대한 구체적으로 설명 (데이터, 수치, 관계, 흐름 등 포함)

한국어로 답변해주세요."""

        request_body = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": mime_type,
                            "data": image_data
                        }
                    }
                ]
            }]
        }

        # API 호출
        response = requests.post(
            f"{GEMINI_API_URL}?key={api_key}",
            headers={"Content-Type": "application/json"},
            json=request_body,
            timeout=30
        )

        response.raise_for_status()
        result = response.json()

        # 응답에서 텍스트 추출
        text = result["candidates"][0]["content"]["parts"][0]["text"]

        elapsed = time.time() - start_time
        log_msg = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 이미지 분석 완료 ({elapsed:.2f}초, {len(image_bytes)} bytes, {mime_type})\n"
        try:
            print(log_msg.strip())
        except UnicodeEncodeError:
            pass  # Windows 콘솔 인코딩 문제 무시
        with open(_get_log_path(), 'a', encoding='utf-8') as f:
            f.write(log_msg)

        return text.strip()

    except Exception as e:
        elapsed = time.time() - start_time
        log_msg = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 이미지 분석 실패 ({elapsed:.2f}초) - {str(e)}\n"
        try:
            print(log_msg.strip())
        except UnicodeEncodeError:
            pass  # Windows 콘솔 인코딩 문제 무시
        with open(_get_log_path(), 'a', encoding='utf-8') as f:
            f.write(log_msg)
            f.write(f"{traceback.format_exc()}\n")
        return None


def get_image_description_markdown(image_bytes: bytes, mime_type: str = "image/png") -> str:
    """
    이미지를 분석하고 마크다운 형식으로 반환

    Returns:
        마크다운 형식의 이미지 + 설명
    """
    # Base64 인라인 이미지
    image_b64 = base64.b64encode(image_bytes).decode('utf-8')
    img_markdown = f"![이미지](data:{mime_type};base64,{image_b64})"

    # 이미지 분석
    description = analyze_image(image_bytes, mime_type)

    if description:
        return f"{img_markdown}\n> **이미지 설명**: {description}\n"
    else:
        return f"{img_markdown}\n"
