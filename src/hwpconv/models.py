"""
데이터 모델 정의

Document, Section, Paragraph, Table, Footnote 등 핵심 데이터 구조
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Union
from enum import Enum


class HeadingLevel(Enum):
    """제목 레벨"""
    NONE = 0
    H1 = 1
    H2 = 2
    H3 = 3
    H4 = 4
    H5 = 5
    H6 = 6


@dataclass
class TextStyle:
    """텍스트 스타일 정보"""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    font_size: Optional[int] = None      # HWPUNIT (10pt = 1000)
    font_name: Optional[str] = None
    color: Optional[str] = None          # "#RRGGBB"
    
    def has_emphasis(self) -> bool:
        """강조 스타일이 있는지 확인"""
        return self.bold or self.italic or self.underline or self.strike


@dataclass
class TextRun:
    """텍스트 런 (동일한 스타일을 가진 텍스트 조각)"""
    text: str
    style: TextStyle = field(default_factory=TextStyle)
    
    def __post_init__(self):
        if self.style is None:
            self.style = TextStyle()


@dataclass
class Paragraph:
    """문단"""
    runs: List[TextRun] = field(default_factory=list)
    heading_level: HeadingLevel = HeadingLevel.NONE
    
    @property
    def text(self) -> str:
        """모든 런의 텍스트를 합친 전체 텍스트"""
        return ''.join(r.text for r in self.runs)
    
    def is_empty(self) -> bool:
        """빈 문단인지 확인"""
        return not self.text.strip()


@dataclass
class TableCell:
    """표 셀"""
    paragraphs: List[Paragraph] = field(default_factory=list)
    rowspan: int = 1
    colspan: int = 1
    
    @property
    def text(self) -> str:
        """셀 내 모든 문단의 텍스트"""
        return '\n'.join(p.text for p in self.paragraphs)
    
    def is_empty(self) -> bool:
        """빈 셀인지 확인"""
        return not self.text.strip()


@dataclass
class TableRow:
    """표 행"""
    cells: List[TableCell] = field(default_factory=list)
    
    @property
    def cell_count(self) -> int:
        """셀 개수"""
        return len(self.cells)


@dataclass
class Table:
    """표"""
    rows: List[TableRow] = field(default_factory=list)
    col_count: int = 0
    
    @property
    def row_count(self) -> int:
        """행 개수"""
        return len(self.rows)
    
    def is_empty(self) -> bool:
        """빈 표인지 확인"""
        return not self.rows


@dataclass
class Footnote:
    """각주/미주"""
    id: str
    number: int
    content: List[Paragraph] = field(default_factory=list)
    
    @property
    def text(self) -> str:
        """각주 내용 텍스트"""
        return '\n'.join(p.text for p in self.content)


@dataclass
class Image:
    """이미지"""
    id: str                           # 이미지 ID
    data: bytes = field(repr=False)   # 이미지 바이너리 데이터
    format: str = 'png'               # 이미지 형식 (png, jpg, gif, etc)
    width: Optional[int] = None       # 너비 (픽셀)
    height: Optional[int] = None      # 높이 (픽셀)
    alt_text: str = ''                # 대체 텍스트
    description: Optional[str] = None # AI 분석 설명
    
    @property
    def base64(self) -> str:
        """Base64 인코딩된 데이터"""
        import base64
        return base64.b64encode(self.data).decode('ascii')
    
    @property
    def data_uri(self) -> str:
        """Data URI 형식 (BMP/TIFF는 PNG로 변환)"""
        import base64
        
        # 지원되지 않는 포맷은 PNG로 변환
        format_lower = self.format.lower()
        if format_lower in ('bmp', 'tif', 'tiff', 'wmf', 'emf'):
            try:
                from PIL import Image as PilImage
                import io
                img = PilImage.open(io.BytesIO(self.data))
                output = io.BytesIO()
                img.save(output, format='PNG')
                converted_data = base64.b64encode(output.getvalue()).decode('ascii')
                return f'data:image/png;base64,{converted_data}'
            except Exception:
                pass  # 변환 실패 시 원본 사용
        
        mime_types = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'gif': 'image/gif',
            'tif': 'image/tiff',
            'tiff': 'image/tiff',
            'webp': 'image/webp',
        }
        mime = mime_types.get(format_lower, 'application/octet-stream')
        return f'data:{mime};base64,{self.base64}'
    
    @property
    def size_kb(self) -> float:
        """파일 크기 (KB)"""
        return len(self.data) / 1024


@dataclass
class Section:
    """구역 (섹션)"""
    elements: List[Union[Paragraph, Table, 'Image']] = field(default_factory=list)
    
    @property
    def paragraph_count(self) -> int:
        """문단 개수"""
        return sum(1 for e in self.elements if isinstance(e, Paragraph))
    
    @property
    def table_count(self) -> int:
        """표 개수"""
        return sum(1 for e in self.elements if isinstance(e, Table))
    
    @property
    def image_count(self) -> int:
        """이미지 개수"""
        return sum(1 for e in self.elements if isinstance(e, Image))


@dataclass
class Document:
    """문서"""
    sections: List[Section] = field(default_factory=list)
    footnotes: Dict[str, Footnote] = field(default_factory=dict)
    endnotes: Dict[str, Footnote] = field(default_factory=dict)
    metadata: Dict[str, str] = field(default_factory=dict)
    images: Dict[str, Image] = field(default_factory=dict)  # id -> Image
    
    @property
    def text(self) -> str:
        """문서 전체 텍스트"""
        texts = []
        for section in self.sections:
            for elem in section.elements:
                if isinstance(elem, Paragraph):
                    texts.append(elem.text)
                elif isinstance(elem, Table):
                    for row in elem.rows:
                        for cell in row.cells:
                            texts.append(cell.text)
        return '\n'.join(texts)
    
    @property
    def section_count(self) -> int:
        """구역 개수"""
        return len(self.sections)
    
    @property
    def total_paragraph_count(self) -> int:
        """전체 문단 개수"""
        return sum(s.paragraph_count for s in self.sections)
    
    @property
    def total_table_count(self) -> int:
        """전체 표 개수"""
        return sum(s.table_count for s in self.sections)
    
    @property
    def total_image_count(self) -> int:
        """전체 이미지 개수"""
        return len(self.images)
