# hwpconv - HWP/HWPX → Markdown 변환기

**목적**: 한글 문서(.hwp, .hwpx)를 Markdown/HTML/Text로 변환하는 Python 라이브러리 및 CLI

---

## 1. 프로젝트 구조

```
hwpconv/
├── pyproject.toml
├── README.md
├── LICENSE                      # MIT
├── src/
│   └── hwpconv/
│       ├── __init__.py
│       ├── cli.py               # CLI 엔트리포인트
│       ├── models.py            # 데이터 모델 (Document, Section, Para, Table, Footnote)
│       ├── parsers/
│       │   ├── __init__.py
│       │   ├── base.py          # BaseParser ABC
│       │   ├── hwpx.py          # HWPX 파서 (ZIP+XML)
│       │   └── hwp.py           # HWP 파서 (OLE+zlib)
│       ├── converters/
│       │   ├── __init__.py
│       │   ├── base.py          # BaseConverter ABC
│       │   ├── markdown.py      # Markdown 변환기
│       │   └── html.py          # HTML 변환기
│       └── utils.py             # 유틸리티
└── tests/
    ├── test_hwpx.py
    ├── test_hwp.py
    └── fixtures/                # 테스트 파일
```

---

## 2. 핵심 데이터 모델 (models.py)

```python
from dataclasses import dataclass, field
from typing import List, Dict, Optional
from enum import Enum

class HeadingLevel(Enum):
    NONE = 0
    H1 = 1
    H2 = 2
    H3 = 3
    H4 = 4

@dataclass
class TextStyle:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    font_size: Optional[int] = None      # HWPUNIT (10pt = 1000)
    font_name: Optional[str] = None
    color: Optional[str] = None          # "#RRGGBB"

@dataclass
class TextRun:
    text: str
    style: TextStyle = field(default_factory=TextStyle)

@dataclass
class Paragraph:
    runs: List[TextRun] = field(default_factory=list)
    heading_level: HeadingLevel = HeadingLevel.NONE
    
    @property
    def text(self) -> str:
        return ''.join(r.text for r in self.runs)

@dataclass
class TableCell:
    paragraphs: List[Paragraph] = field(default_factory=list)
    rowspan: int = 1
    colspan: int = 1
    
    @property
    def text(self) -> str:
        return '\n'.join(p.text for p in self.paragraphs)

@dataclass
class TableRow:
    cells: List[TableCell] = field(default_factory=list)

@dataclass
class Table:
    rows: List[TableRow] = field(default_factory=list)
    col_count: int = 0

@dataclass
class Footnote:
    id: str
    number: int
    content: List[Paragraph] = field(default_factory=list)
    
    @property
    def text(self) -> str:
        return '\n'.join(p.text for p in self.content)

@dataclass
class Section:
    elements: List = field(default_factory=list)  # Paragraph | Table

@dataclass
class Document:
    sections: List[Section] = field(default_factory=list)
    footnotes: Dict[str, Footnote] = field(default_factory=dict)
    endnotes: Dict[str, Footnote] = field(default_factory=dict)
    metadata: Dict[str, str] = field(default_factory=dict)
```

---

## 3. HWPX 파서 (parsers/hwpx.py)

### 3.0 Base Parser (parsers/base.py)

```python
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Union
from ..models import Document

class BaseParser(ABC):
    """파서 베이스 클래스"""
    
    @abstractmethod
    def parse(self, file_path: str) -> Document:
        """파싱 실행"""
        pass
    
    @staticmethod
    @abstractmethod
    def quick_extract(file_path: str) -> str:
        """빠른 텍스트 추출 (Preview 활용)"""
        pass
    
    @classmethod
    def can_parse(cls, file_path: str) -> bool:
        """파싱 가능 여부"""
        return Path(file_path).suffix.lower() in cls.SUPPORTED_EXTENSIONS


class BaseConverter(ABC):
    """변환기 베이스 클래스"""
    
    @abstractmethod
    def convert(self, doc: Document) -> str:
        """변환 실행"""
        pass
    
    def save(self, doc: Document, output_path: str) -> None:
        """파일로 저장"""
        content = self.convert(doc)
        Path(output_path).write_text(content, encoding='utf-8')
```

## 3. HWPX 파서 (parsers/hwpx.py)

### 3.1 파일 구조

```
HWPX (ZIP)
├── mimetype                # "application/hwp+zip"
├── version.xml             # OWPML 버전 정보
├── Contents/
│   ├── header.xml          # 서식 정보 (charPr, paraPr, fontface) ★
│   ├── section0.xml        # 본문 (p → run → t, tbl, footNote) ★
│   ├── section1.xml...     # 추가 구역
│   └── content.hpf         # OPF 패키징 정보
├── Preview/
│   └── PrvText.txt         # 빠른 텍스트 추출용 ★
└── BinData/                # 이미지 (image*.png 등)
```

### 3.2 XML 네임스페이스

```python
NS = {
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',      # header.xml
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph', # section.xml 본문
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',   # section.xml 루트
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',      # 공통
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',       # 앱 정보
    'hm': 'http://www.hancom.co.kr/hwpml/2011/master',    # 마스터 페이지
    'hv': 'http://www.hancom.co.kr/hwpml/2011/version',   # 버전
}
```

### 3.3 header.xml 구조 (서식 정의)

```xml
<hh:head>
  <hh:refList>
    <!-- 글꼴 정의 -->
    <hh:fontfaces>
      <hh:fontface lang="HANGUL" fontCnt="2">
        <hh:font id="0" face="함초롬바탕" type="TTF"/>
      </hh:fontface>
    </hh:fontfaces>
    
    <!-- 글자 모양 정의 -->
    <hh:charProperties itemCnt="10">
      <hh:charPr id="0" height="1000" textColor="#000000">
        <hh:fontRef hangul="0" latin="0"/>
      </hh:charPr>
      <hh:charPr id="1" height="1400" textColor="#000000" bold="true"/>
    </hh:charProperties>
    
    <!-- 문단 모양 정의 -->
    <hh:paraProperties itemCnt="5">
      <hh:paraPr id="0" align="JUSTIFY">
        <hh:margin left="0" right="0"/>
      </hh:paraPr>
    </hh:paraProperties>
  </hh:refList>
</hh:head>
```

**charPr 주요 속성**:
| 속성 | 설명 | 예시 |
|------|------|------|
| id | 참조 ID | "0", "1" |
| height | 글자 크기 (HWPUNIT, 10pt=1000) | "1000" |
| textColor | 글자 색상 | "#000000" |
| bold | 굵게 | "true" |
| italic | 기울임 | "true" |
| underline | 밑줄 | "true" |
| strikeout | 취소선 | "true" |

### 3.4 section.xml 구조 (본문)

```xml
<hs:sec>
  <!-- 문단 -->
  <hp:p paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>일반 텍스트</hp:t>
    </hp:run>
    <hp:run charPrIDRef="1">
      <hp:t>굵은 텍스트</hp:t>
    </hp:run>
  </hp:p>
  
  <!-- 표 -->
  <hp:p>
    <hp:run>
      <hp:tbl colCnt="3" rowCnt="2">
        <hp:tr>
          <hp:tc colSpan="1" rowSpan="1">
            <hp:p><hp:run><hp:t>셀1</hp:t></hp:run></hp:p>
          </hp:tc>
          <hp:tc>...</hp:tc>
        </hp:tr>
      </hp:tbl>
    </hp:run>
  </hp:p>
  
  <!-- 각주 포함 문단 -->
  <hp:p>
    <hp:run>
      <hp:t>본문 텍스트</hp:t>
      <hp:footNote id="fn1">
        <hp:p><hp:run><hp:t>각주 내용</hp:t></hp:run></hp:p>
      </hp:footNote>
    </hp:run>
  </hp:p>
</hs:sec>
```

**hp:t 요소 (mixed content)**:
```xml
<hp:t>첫 번째 줄<lineBreak/>두 번째 줄<tab/>탭 뒤 텍스트</hp:t>
```
→ `첫 번째 줄\n두 번째 줄\t탭 뒤 텍스트`

### 3.5 구현

```python
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional
from .base import BaseParser
from ..models import (
    Document, Section, Paragraph, TextRun, Table, TableRow, TableCell,
    TextStyle, Footnote
)

class HwpxParser(BaseParser):
    NS = {
        'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
        'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    }
    
    def __init__(self):
        self.char_shapes: Dict[str, TextStyle] = {}  # id -> TextStyle
        self.para_shapes: Dict[str, dict] = {}       # id -> {align, ...}
    
    def parse(self, file_path: str) -> Document:
        doc = Document()
        
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 1. 네임스페이스 추출 (동적)
            self._extract_namespaces(zf)
            
            # 2. header.xml에서 서식 정보 로드
            self._load_header(zf)
            
            # 3. 구역 파일 목록
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])
            
            # 4. 각 section 파싱
            for sf in section_files:
                section = self._parse_section(zf, sf)
                doc.sections.append(section)
        
        return doc
    
    def _extract_namespaces(self, zf: zipfile.ZipFile):
        """header.xml에서 실제 네임스페이스 추출"""
        if 'Contents/header.xml' not in zf.namelist():
            return
        
        xml_data = zf.read('Contents/header.xml')
        for event, elem in ET.iterparse(
            __import__('io').BytesIO(xml_data), events=['start-ns']
        ):
            prefix, uri = elem
            if prefix:
                self.NS[prefix] = uri
    
    def _load_header(self, zf: zipfile.ZipFile):
        """header.xml에서 charPr, paraPr 로드"""
        if 'Contents/header.xml' not in zf.namelist():
            return
        
        xml_data = zf.read('Contents/header.xml')
        root = ET.fromstring(xml_data)
        
        # refList 찾기
        ref_list = root.find('.//hh:refList', self.NS)
        if ref_list is None:
            # 네임스페이스 없이 시도
            ref_list = root.find('.//{*}refList')
        
        if ref_list is None:
            return
        
        # 글자 모양 (charPr)
        char_props = ref_list.find('.//hh:charProperties', self.NS)
        if char_props is None:
            char_props = ref_list.find('.//{*}charProperties')
        
        if char_props:
            for cp in char_props.findall('.//{*}charPr'):
                cp_id = cp.get('id', '0')
                style = TextStyle()
                
                # 속성 파싱
                style.font_size = int(cp.get('height', 1000))
                style.color = cp.get('textColor', '#000000')
                style.bold = cp.get('bold', 'false').lower() == 'true'
                style.italic = cp.get('italic', 'false').lower() == 'true'
                style.underline = cp.get('underline', 'false').lower() == 'true'
                style.strike = cp.get('strikeout', 'false').lower() == 'true'
                
                self.char_shapes[cp_id] = style
    
    def _parse_section(self, zf: zipfile.ZipFile, path: str) -> Section:
        """section*.xml 파싱"""
        section = Section()
        xml_data = zf.read(path)
        root = ET.fromstring(xml_data)
        
        # 모든 p 요소 순회
        for p_elem in root.iter():
            local_name = p_elem.tag.split('}')[-1] if '}' in p_elem.tag else p_elem.tag
            
            if local_name == 'p':
                # 표가 포함된 문단인지 확인
                tbl = self._find_child(p_elem, 'tbl')
                if tbl is not None:
                    table = self._parse_table(tbl)
                    section.elements.append(table)
                else:
                    para = self._parse_paragraph(p_elem)
                    if para.text.strip():
                        section.elements.append(para)
        
        return section
    
    def _parse_paragraph(self, p_elem) -> Paragraph:
        """p 요소 파싱"""
        para = Paragraph()
        
        for run in self._find_all_children(p_elem, 'run'):
            char_pr_id = run.get('charPrIDRef', '0')
            style = self.char_shapes.get(char_pr_id, TextStyle())
            
            for child in run:
                local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                if local_name == 't':
                    text = self._extract_text(child)
                    if text:
                        para.runs.append(TextRun(text=text, style=style))
        
        return para
    
    def _extract_text(self, t_elem) -> str:
        """t 요소에서 텍스트 추출 (mixed content 처리)"""
        parts = []
        
        # t.text: 첫 번째 텍스트
        if t_elem.text:
            parts.append(t_elem.text)
        
        # 자식 요소 처리
        for child in t_elem:
            local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if local_name == 'lineBreak':
                parts.append('\n')
            elif local_name == 'tab':
                parts.append('\t')
            
            # child.tail: 자식 요소 뒤의 텍스트
            if child.tail:
                parts.append(child.tail)
        
        return ''.join(parts)
    
    def _parse_table(self, tbl_elem) -> Table:
        """tbl 요소 파싱"""
        table = Table()
        table.col_count = int(tbl_elem.get('colCnt', 0))
        
        for tr in self._find_all_children(tbl_elem, 'tr'):
            row = TableRow()
            
            for tc in self._find_all_children(tr, 'tc'):
                cell = TableCell()
                cell.colspan = int(tc.get('colSpan', 1))
                cell.rowspan = int(tc.get('rowSpan', 1))
                
                # 셀 내 문단들
                for p in self._find_all_children(tc, 'p'):
                    para = self._parse_paragraph(p)
                    cell.paragraphs.append(para)
                
                row.cells.append(cell)
            
            table.rows.append(row)
        
        return table
    
    def _find_child(self, elem, local_name: str):
        """로컬 이름으로 자식 요소 찾기"""
        for child in elem:
            child_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if child_name == local_name:
                return child
        return None
    
    def _find_all_children(self, elem, local_name: str):
        """로컬 이름으로 모든 자식 요소 찾기"""
        result = []
        for child in elem:
            child_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if child_name == local_name:
                result.append(child)
        return result
    
    @staticmethod
    def quick_extract(file_path: str) -> str:
        """PrvText.txt에서 빠른 텍스트 추출"""
        with zipfile.ZipFile(file_path, 'r') as zf:
            if 'Preview/PrvText.txt' in zf.namelist():
                return zf.read('Preview/PrvText.txt').decode('utf-8', errors='ignore')
        return ''
```

---

## 4. HWP 파서 (parsers/hwp.py)

### 4.1 파일 구조

```
HWP (OLE Compound Document)
├── FileHeader          # 256 bytes, 시그니처 "HWP Document File"
├── DocInfo             # 문서 정보 (zlib 압축) - 글꼴, 스타일 정의
├── BodyText/
│   ├── Section0        # 본문 (zlib 압축, 레코드 구조)
│   └── Section1...
├── BinData/            # 이미지
└── PrvText             # 미리보기 텍스트 (UTF-16LE) ★ 빠른 추출용
```

### 4.2 레코드 구조 (Little-Endian)

```
┌─────────────────────────────────────┐
│ Record Header (4 bytes)             │
│ ├─ TagID: bits 0-9 (10 bits)        │
│ ├─ Level: bits 10-19 (10 bits)      │
│ └─ Size: bits 20-31 (12 bits)       │  ← 0xFFF면 확장
├─────────────────────────────────────┤
│ [Extended Size] (4 bytes, optional) │  ← Size==0xFFF일 때만
├─────────────────────────────────────┤
│ Data (Size bytes)                   │
└─────────────────────────────────────┘
```

**레코드 헤더 파싱**:
```python
header = struct.unpack('<I', data[pos:pos+4])[0]
tag_id = (header >> 0) & 0x3FF    # 10 bits
level = (header >> 10) & 0x3FF    # 10 bits
size = (header >> 20) & 0xFFF     # 12 bits

if size == 0xFFF:  # 확장 크기
    size = struct.unpack('<I', data[pos+4:pos+8])[0]
```

### 4.3 주요 TagID

**DocInfo 레코드**:
| TagID | 값 | 이름 | 설명 |
|-------|-----|------|------|
| 16 | 0x010 | DOCUMENT_PROPERTIES | 구역 개수, 시작번호 |
| 17 | 0x011 | ID_MAPPINGS | BinData, 글꼴 개수 |
| 18 | 0x012 | BIN_DATA | 이미지 참조 정보 |
| 19 | 0x013 | FACE_NAME | 글꼴 정보 |
| 21 | 0x015 | CHAR_SHAPE | 글자 모양 |
| 22 | 0x016 | PARA_SHAPE | 문단 모양 |

**BodyText 레코드**:
| TagID | 값 | 이름 | 설명 |
|-------|-----|------|------|
| 66 | 0x042 | PARA_HEADER | 문단 헤더 |
| 67 | 0x043 | PARA_TEXT | 문단 텍스트 |
| 68 | 0x044 | PARA_CHAR_SHAPE | 글자 모양 적용 위치 |
| 69 | 0x045 | PARA_LINE_SEG | 줄 레이아웃 |
| 71 | 0x047 | CTRL_HEADER | 컨트롤 헤더 |
| 74 | 0x04A | FOOTNOTE_SHAPE | 각주 모양 |
| 77 | 0x04D | TABLE | 표 |

### 4.4 제어문자 (PARA_TEXT 내)

**제어문자 코드 테이블** (공식 문서 [표 6]):
| 코드 | 의미 | 처리 |
|------|------|------|
| 0 | 문단 끝 | 파싱 종료 |
| 1 | 예약 | +14바이트 스킵 (확장 컨트롤) |
| 2 | 구역/단 정의 | +14바이트 스킵, CTRL_HEADER 참조 |
| 3 | 필드 시작 | +14바이트 스킵 |
| 4-9 | 예약 | 문자 크기 (2바이트) |
| **10** | **줄바꿈** | `\n` 출력 (일반 문자) |
| **11** | **각주/미주** | +14바이트 스킵, 각주 처리 |
| **12** | **그림/표** | +14바이트 스킵, 개체 처리 |
| 13-15 | 예약 | +14바이트 스킵 |
| 16-31 | 예약 | 문자 크기 (2바이트) |

**확장 컨트롤 ID** (14바이트 추가 데이터 내, 역순 저장):
| ID (역순) | 원래 | 의미 |
|-----------|------|------|
| `dces` | secd | 구역 정의 |
| `dloc` | cold | 단 정의 |
| ` lbt` | tbl  | 표 |
| `$cip` | pic$ | 그림 |
| `  nf` | fn   | 각주 |
| `  ne` | en   | 미주 |

### 4.5 PARA_HEADER 구조

```python
# PARA_HEADER 데이터 (최소 22바이트)
char_count = struct.unpack('<I', data[0:4])[0]    # 문자 수 (최상위 비트=마지막 문단)
control_mask = struct.unpack('<I', data[4:8])[0]  # 컨트롤 마스크
para_shape_id = struct.unpack('<H', data[8:10])[0]  # 문단 모양 ID
para_style_id = data[10]                           # 문단 스타일 ID
# ...

# 마지막 문단 체크
is_last = bool(char_count & 0x80000000)
char_count = char_count & 0x7FFFFFFF
```

### 4.6 구현

```python
import olefile
import zlib
import struct
from typing import List, Dict
from .base import BaseParser
from ..models import Document, Section, Paragraph, TextRun, Table, TextStyle

class HwpParser(BaseParser):
    # 레코드 TagID
    HWPTAG_PARA_HEADER = 0x042
    HWPTAG_PARA_TEXT = 0x043
    HWPTAG_CTRL_HEADER = 0x047
    
    # 제어문자
    CTRL_SECTION_COLUMN = 2
    CTRL_FIELD_START = 3
    CTRL_LINE_BREAK = 10
    CTRL_FOOTNOTE = 11
    CTRL_PICTURE_TABLE = 12
    
    # 확장 컨트롤 (14바이트 스킵 필요)
    EXTENDED_CTRLS = {1, 2, 3, 11, 12, 13, 14, 15}
    
    def __init__(self):
        self.char_shapes: Dict[int, TextStyle] = {}
        self.para_shapes: Dict[int, dict] = {}
    
    def parse(self, file_path: str) -> Document:
        doc = Document()
        ole = olefile.OleFileIO(file_path)
        
        try:
            # 1. DocInfo에서 서식 정보 로드 (선택)
            self._load_doc_info(ole)
            
            # 2. Section 스트림들 파싱
            for entry in ole.listdir():
                path = '/'.join(entry)
                if path.startswith('BodyText/Section'):
                    data = ole.openstream(path).read()
                    data = zlib.decompress(data, -15)  # raw deflate
                    
                    section = self._parse_section(data)
                    doc.sections.append(section)
        finally:
            ole.close()
        
        return doc
    
    def _load_doc_info(self, ole):
        """DocInfo에서 CharShape, ParaShape 로드"""
        if not ole.exists('DocInfo'):
            return
        
        data = ole.openstream('DocInfo').read()
        try:
            data = zlib.decompress(data, -15)
        except:
            return  # 압축 안 된 경우
        
        pos = 0
        while pos < len(data) - 4:
            header = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
            
            tag_id = header & 0x3FF
            size = (header >> 20) & 0xFFF
            
            if size == 0xFFF:
                size = struct.unpack('<I', data[pos:pos+4])[0]
                pos += 4
            
            record_data = data[pos:pos+size]
            pos += size
            
            # CHAR_SHAPE (21) 파싱
            if tag_id == 0x015 and len(record_data) >= 72:
                # face_id: 7 WORD (언어별 글꼴)
                # height: HWPUNIT (offset 56)
                # attributes: DWORD (offset 60) - bit0=italic, bit1=bold
                height = struct.unpack('<I', record_data[56:60])[0]
                attrs = struct.unpack('<I', record_data[60:64])[0]
                
                style = TextStyle()
                style.font_size = height
                style.italic = bool(attrs & 0x01)
                style.bold = bool(attrs & 0x02)
                
                shape_id = len(self.char_shapes)
                self.char_shapes[shape_id] = style
    
    def _parse_section(self, data: bytes) -> Section:
        """Section 데이터 파싱"""
        section = Section()
        pos = 0
        
        while pos < len(data) - 4:
            # 레코드 헤더 읽기
            header = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
            
            tag_id = header & 0x3FF
            size = (header >> 20) & 0xFFF
            
            if size == 0xFFF:
                size = struct.unpack('<I', data[pos:pos+4])[0]
                pos += 4
            
            record_data = data[pos:pos+size]
            pos += size
            
            # PARA_TEXT 처리
            if tag_id == self.HWPTAG_PARA_TEXT:
                text = self._extract_text(record_data)
                if text.strip():
                    para = Paragraph()
                    para.runs.append(TextRun(text=text))
                    section.elements.append(para)
        
        return section
    
    def _extract_text(self, data: bytes) -> str:
        """PARA_TEXT 레코드에서 텍스트 추출"""
        chars = []
        i = 0
        
        while i < len(data) - 1:
            code = struct.unpack('<H', data[i:i+2])[0]
            i += 2
            
            # 제어문자 처리
            if code < 32:
                if code == 0:  # 문단 끝
                    break
                elif code == self.CTRL_LINE_BREAK:  # 줄바꿸
                    chars.append('\n')
                elif code in self.EXTENDED_CTRLS:
                    # 확장 컨트롤 - 14바이트 추가 데이터 스킵
                    i += 14
                continue
            
            # 일반 문자 (UTF-16LE)
            chars.append(chr(code))
        
        return ''.join(chars)
    
    @staticmethod
    def quick_extract(file_path: str) -> str:
        """PrvText에서 빠른 텍스트 추출"""
        ole = olefile.OleFileIO(file_path)
        try:
            if ole.exists('PrvText'):
                data = ole.openstream('PrvText').read()
                return data.decode('utf-16-le', errors='ignore')
            return ''
        finally:
            ole.close()
```

---

## 5. Markdown 변환기 (converters/markdown.py)

```python
from typing import List
from .base import BaseConverter
from ..models import Document, Section, Paragraph, Table, TextRun, Footnote, HeadingLevel

class MarkdownConverter(BaseConverter):
    def convert(self, doc: Document) -> str:
        lines = []
        footnote_counter = 1
        
        for section in doc.sections:
            for elem in section.elements:
                if isinstance(elem, Paragraph):
                    line = self._convert_paragraph(elem)
                    if line:
                        lines.append(line)
                        lines.append('')  # 빈 줄
                elif isinstance(elem, Table):
                    table_md = self._convert_table(elem)
                    lines.append(table_md)
                    lines.append('')
        
        # 각주 추가
        if doc.footnotes:
            lines.append('')
            lines.append('---')
            lines.append('')
            for fn_id, fn in sorted(doc.footnotes.items(), key=lambda x: x[1].number):
                lines.append(f'[^{fn.number}]: {fn.text}')
        
        return '\n'.join(lines)
    
    def _convert_paragraph(self, para: Paragraph) -> str:
        """문단 → Markdown"""
        text = ''
        
        for run in para.runs:
            run_text = run.text
            
            # 스타일 적용
            if run.style.bold and run.style.italic:
                run_text = f'***{run_text}***'
            elif run.style.bold:
                run_text = f'**{run_text}**'
            elif run.style.italic:
                run_text = f'*{run_text}*'
            
            if run.style.strike:
                run_text = f'~~{run_text}~~'
            
            text += run_text
        
        # 제목 레벨
        if para.heading_level != HeadingLevel.NONE:
            prefix = '#' * para.heading_level.value
            text = f'{prefix} {text}'
        
        return text
    
    def _convert_table(self, table: Table) -> str:
        """표 → Markdown"""
        if not table.rows:
            return ''
        
        lines = []
        col_count = table.col_count or len(table.rows[0].cells)
        
        for i, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                cell_text = cell.text.replace('\n', ' ').replace('|', '\\|')
                cells.append(cell_text)
            
            # 컬럼 수 맞추기
            while len(cells) < col_count:
                cells.append('')
            
            lines.append('| ' + ' | '.join(cells) + ' |')
            
            # 헤더 구분선 (첫 행 다음)
            if i == 0:
                lines.append('| ' + ' | '.join(['---'] * col_count) + ' |')
        
        return '\n'.join(lines)
```

---

## 6. CLI (cli.py)

```python
import argparse
import sys
from pathlib import Path
from .parsers.hwpx import HwpxParser
from .parsers.hwp import HwpParser
from .converters.markdown import MarkdownConverter

def main():
    parser = argparse.ArgumentParser(
        prog='hwpconv',
        description='HWP/HWPX → Markdown/HTML 변환기'
    )
    parser.add_argument('input', help='입력 파일 (.hwp, .hwpx)')
    parser.add_argument('-o', '--output', help='출력 파일')
    parser.add_argument('-f', '--format', choices=['md', 'html', 'txt'], 
                        default='md', help='출력 포맷')
    parser.add_argument('--quick', action='store_true', 
                        help='빠른 텍스트 추출 (HWP only)')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f'Error: {input_path} not found', file=sys.stderr)
        sys.exit(1)
    
    # 파서 선택
    ext = input_path.suffix.lower()
    if ext == '.hwpx':
        doc = HwpxParser().parse(str(input_path))
    elif ext == '.hwp':
        if args.quick:
            result = HwpParser.quick_extract(str(input_path))
            print(result)
            return
        doc = HwpParser().parse(str(input_path))
    else:
        print(f'Error: Unsupported format {ext}', file=sys.stderr)
        sys.exit(1)
    
    # 변환
    result = MarkdownConverter().convert(doc)
    
    # 출력
    if args.output:
        Path(args.output).write_text(result, encoding='utf-8')
        print(f'Saved to {args.output}')
    else:
        print(result)

if __name__ == '__main__':
    main()
```

---

## 7. pyproject.toml

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "hwpconv"
version = "0.1.0"
description = "HWP/HWPX → Markdown/HTML 변환기"
readme = "README.md"
license = "MIT"
requires-python = ">=3.9"
authors = [
    { name = "Renaissance Law Firm", email = "info@rnlaw.co.kr" }
]
keywords = ["hwp", "hwpx", "markdown", "converter", "한글"]
classifiers = [
    "Development Status :: 3 - Alpha",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.9",
    "Programming Language :: Python :: 3.10",
    "Programming Language :: Python :: 3.11",
    "Programming Language :: Python :: 3.12",
]
dependencies = [
    "olefile>=0.46",
]

[project.optional-dependencies]
dev = [
    "pytest>=7.0",
    "pytest-cov",
]

[project.scripts]
hwpconv = "hwpconv.cli:main"

[project.urls]
Homepage = "https://github.com/rnlaw/hwpconv"

[tool.hatch.build.targets.wheel]
packages = ["src/hwpconv"]
```

---

## 8. 구현 우선순위

### Phase 1: MVP (4-6시간)
- [x] 프로젝트 구조
- [ ] models.py - 데이터 모델
- [ ] HwpxParser - 텍스트 추출
- [ ] HwpParser - PrvText 빠른 추출
- [ ] MarkdownConverter - 기본 변환
- [ ] CLI

### Phase 2: 표 지원 (3-4시간)
- [ ] HwpxParser - 표 파싱
- [ ] HwpParser - CTRL_HEADER 표 처리
- [ ] MarkdownConverter - 표 출력

### Phase 3: 서식 (2-3시간)
- [ ] HwpxParser - charPr 파싱 (bold, italic)
- [ ] HwpParser - CharShape 파싱
- [ ] 제목 감지 휴리스틱

### Phase 4: 각주 (1-2시간)
- [ ] HwpxParser - footNote 파싱
- [ ] HwpParser - 제어문자 11 처리
- [ ] Markdown [^n] 출력

---

## 9. 사용 예시

```bash
# 설치
pip install hwpconv

# CLI 사용
hwpconv document.hwpx -o output.md
hwpconv document.hwp -o output.md
hwpconv document.hwp --quick  # 빠른 텍스트만

# Python API
from hwpconv import HwpxParser, MarkdownConverter

doc = HwpxParser().parse('document.hwpx')
md = MarkdownConverter().convert(doc)
print(md)
```

---

## 10. 실제 테스트 결과 (검증 완료)

### 10.1 테스트 파일

| 파일 | 크기 | 형식 | 특징 |
|------|------|------|------|
| 회생_기각결정문_반박내용.hwpx | 231KB | HWPX | 텍스트 위주, 표 없음 |
| 비영리법인_설립신청_안내.hwp | 154KB | HWP | 표 29개, 이미지 없음 |
| 정관_작성기준_및_예시.hwp | 41KB | HWP | 표 2개, 이미지 없음 |

### 10.2 HWPX 테스트 결과

```
회생_기각결정문_반박내용.hwpx (231KB)
├── Contents/section0.xml: 109KB
├── Preview/PrvText.txt: 2.5KB
└── 추출 결과: 7개 문단, 텍스트 100% 보존
```

### 10.3 HWP 테스트 결과

```
비영리법인_설립신청_안내.hwp (154KB)
├── BodyText/Section0: 60KB → 258KB (zlib)
├── PrvText: UTF-16LE 텍스트
├── HWPTAG_TABLE 레코드: 29개
└── 추출 결과: PrvText 빠른 추출 성공
```

---

## 11. 참고 자료

### 11.1 한컴 공식 리소스

| 리소스 | URL |
|--------|-----|
| **스펙 다운로드** | https://www.hancom.com/support/downloadCenter/hwpOwpml |
| **GitHub** | https://github.com/hancom-io |
| **개발자 포털** | https://developer.hancom.com |
| **개발자 포럼** | https://forum.developer.hancom.com |

### 11.2 공식 스펙 문서

| 문서 | 설명 |
|------|------|
| **HWP 5.0 형식** | 바이너리 포맷 (2002~현재) - 핵심 |
| 배포용 문서 형식 | hwpx 배포 포맷 |
| 수식 형식 | 수식 개체 스펙 |
| 차트 형식 | 차트 개체 스펙 |

### 11.3 한컴테크 블로그 (공식)

| 제목 | 저자 |
|------|------|
| HWP 포맷 구조 살펴보기 | 정우진 |
| HWPX 포맷 구조 살펴보기 | 김규리 |
| Python을 통한 HWP 파싱 (1), (2) | 정우진 |
| Python을 통한 HWPX 파싱 (1), (2) | 김규리 |

URL: https://tech.hancom.com

### 11.4 오픈소스 참고

| 이름 | 언어 | 라이선스 |
|------|------|----------|
| pyhwp | Python | AGPL-3.0 |
| pyhwpx | Python | MIT |
| file2md | Node.js | MIT |
| hanpama/hwp | Go | - |

### 11.5 저작권 고지 (필수)

개발 결과물에 반드시 포함:
```
본 제품은 한글과컴퓨터의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
```

---

## 12. 라이선스

MIT License
