"""
변환기 베이스 클래스
"""

from abc import ABC, abstractmethod
from pathlib import Path

from ..models import Document


class BaseConverter(ABC):
    """변환기 베이스 클래스"""
    
    @abstractmethod
    def convert(self, doc: Document) -> str:
        """Document를 문자열로 변환
        
        Args:
            doc: 변환할 Document 객체
            
        Returns:
            str: 변환된 문자열
        """
        pass
    
    def save(self, doc: Document, output_path: str) -> None:
        """Document를 파일로 저장
        
        Args:
            doc: 변환할 Document 객체
            output_path: 출력 파일 경로
        """
        content = self.convert(doc)
        Path(output_path).write_text(content, encoding='utf-8')
