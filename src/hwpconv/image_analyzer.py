"""
Gemini Vision API를 사용한 이미지 분석 모듈
"""
import base64
from typing import Optional
import traceback

try:
    import google.generativeai as genai
except ImportError:
    genai = None

from . import config as app_config

# 지원되는 MIME 타입 (Gemini API)
SUPPORTED_MIME_TYPES = {'image/png', 'image/jpeg', 'image/webp', 'image/heic', 'image/heif'}

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
    return genai is not None and app_config.has_api_key()


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
        genai.configure(api_key=api_key)
        
        model = genai.GenerativeModel('gemini-3-flash-preview')
        
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
        
        # API 호출
        response = model.generate_content([
            "이 이미지의 내용을 한국어로 간단히 설명해주세요 (1-2문장):",
            {
                "mime_type": mime_type,
                "data": image_data
            }
        ])
        
        elapsed = time.time() - start_time
        log_msg = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 이미지 분석 완료 ({elapsed:.2f}초, {len(image_bytes)} bytes, {mime_type})\n"
        try:
            print(log_msg.strip())
        except UnicodeEncodeError:
            pass  # Windows 콘솔 인코딩 문제 무시
        with open(_get_log_path(), 'a', encoding='utf-8') as f:
            f.write(log_msg)

        return response.text.strip()

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
        return f"{img_markdown}\n> 🖼️ **이미지 설명**: {description}\n"
    else:
        return f"{img_markdown}\n"
