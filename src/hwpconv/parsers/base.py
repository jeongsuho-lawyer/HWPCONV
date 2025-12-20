"""
파서 베이스 클래스
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Set

from ..models import Document


class BaseParser(ABC):
    """파서 베이스 클래스"""
    
    SUPPORTED_EXTENSIONS: Set[str] = set()
    
    @abstractmethod
    def parse(self, file_path: str) -> Document:
        """파일을 파싱하여 Document 객체 반환
        
        Args:
            file_path: 파싱할 파일 경로
            
        Returns:
            Document: 파싱된 문서 객체
        """
        pass
    
    @staticmethod
    @abstractmethod
    def quick_extract(file_path: str) -> str:
        """빠른 텍스트 추출 (Preview 활용)
        
        Args:
            file_path: 파일 경로
            
        Returns:
            str: 추출된 텍스트
        """
        pass
    
    @classmethod
    def can_parse(cls, file_path: str) -> bool:
        """파싱 가능 여부 확인
        
        Args:
            file_path: 파일 경로
            
        Returns:
            bool: 파싱 가능 여부
        """
        return Path(file_path).suffix.lower() in cls.SUPPORTED_EXTENSIONS
