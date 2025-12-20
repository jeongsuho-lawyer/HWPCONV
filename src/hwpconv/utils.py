"""
유틸리티 함수
"""

from pathlib import Path
from typing import Union


def detect_format(file_path: Union[str, Path]) -> str:
    """파일 형식 감지
    
    Args:
        file_path: 파일 경로
        
    Returns:
        str: 'hwp', 'hwpx', 또는 'unknown'
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    
    if ext == '.hwp':
        return 'hwp'
    elif ext == '.hwpx':
        return 'hwpx'
    
    # 확장자가 없거나 다른 경우 파일 시그니처 확인
    try:
        with open(path, 'rb') as f:
            header = f.read(8)
            
            # ZIP 시그니처 (HWPX)
            if header[:4] == b'PK\x03\x04':
                return 'hwpx'
            
            # OLE 시그니처 (HWP)
            if header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                return 'hwp'
    except Exception:
        pass
    
    return 'unknown'


def hwpunit_to_pt(hwpunit: int) -> float:
    """HWPUNIT을 포인트로 변환
    
    Args:
        hwpunit: HWPUNIT 값 (1000 = 10pt)
        
    Returns:
        float: 포인트 값
    """
    return hwpunit / 100.0


def pt_to_hwpunit(pt: float) -> int:
    """포인트를 HWPUNIT으로 변환
    
    Args:
        pt: 포인트 값
        
    Returns:
        int: HWPUNIT 값
    """
    return int(pt * 100)
