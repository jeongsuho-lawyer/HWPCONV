"""
HWP 파서

HWP (OLE Compound Document) 형식 파일을 파싱하는 모듈
"""

import struct
import zlib
from typing import Dict, List, Optional, Set, Tuple

import olefile

from .base import BaseParser
from ..models import (
    Document, Section, Paragraph, TextRun, Table, TableRow, TableCell,
    TextStyle, Footnote, HeadingLevel, Image
)


class HwpParser(BaseParser):
    """HWP (OLE Compound Document) 파일 파서"""
    
    SUPPORTED_EXTENSIONS: Set[str] = {'.hwp'}
    
    # DocInfo 레코드 TagID
    HWPTAG_DOCUMENT_PROPERTIES = 0x010  # 16
    HWPTAG_ID_MAPPINGS = 0x011          # 17
    HWPTAG_BIN_DATA = 0x012             # 18
    HWPTAG_FACE_NAME = 0x013            # 19
    HWPTAG_CHAR_SHAPE = 0x015           # 21
    HWPTAG_PARA_SHAPE = 0x016           # 22
    
    # BodyText 레코드 TagID (Standard HWP 5.0)
    HWPTAG_PARA_HEADER = 16
    HWPTAG_PARA_TEXT = 17
    HWPTAG_PARA_CHAR_SHAPE = 18
    HWPTAG_PARA_LINE_SEG = 19
    HWPTAG_CTRL_HEADER = 21
    HWPTAG_LIST_HEADER = 24  # Standard
    HWPTAG_TABLE = 27        # 77 - 50 (normalized from shifted tag)
    HWPTAG_FOOTNOTE_SHAPE = 74 # ?
    
    # Shifted tags (+50)
    TAG_SHIFT_OFFSET = 50
    SHIFTED_PARA_HEADER = 66     # 16 + 50
    SHIFTED_PARA_TEXT = 67       # 17 + 50
    SHIFTED_PARA_CHAR_SHAPE = 68 # 18 + 50
    SHIFTED_LIST_HEADER = 72     # 22 + 50 (실제로는 24+48=72)
    
    # 제어문자 타입
    CTRL_CHAR_TYPE = {0, 10, 13, 24, 30, 31}  # 1 WCHAR
    CTRL_INLINE_TYPE = {4, 9}  # 8 WCHAR, 위치 1개 차지
    CTRL_EXTENDED_TYPE = {2, 3, 11, 15, 16, 17, 18, 21, 22, 23}  # 8 WCHAR
    
    # 컨트롤 ID (4-byte)
    CTRL_ID_TABLE = 0x74626C20  # 'tbl '
    CTRL_ID_PICTURE = 0x24706963  # '$pic'
    CTRL_ID_OLE = 0x246F6C65  # '$ole'
    CTRL_ID_EQUATION = 0x65716564  # 'eqed'
    
    # 제어문자 코드
    CTRL_PARA_END = 0           # 문단 끝
    CTRL_RESERVED_1 = 1         # 예약 (+14바이트)
    CTRL_SECTION_COLUMN = 2     # 구역/단 정의 (+14바이트)
    CTRL_FIELD_START = 3        # 필드 시작 (+14바이트)
    CTRL_LINE_BREAK = 10        # 줄바꿈 (일반 문자 크기)
    CTRL_FOOTNOTE = 11          # 각주/미주 (+14바이트)
    CTRL_PICTURE_TABLE = 12     # 그림/표 (+14바이트)
    
    # 확장 컨트롤 (14바이트 스킵 필요)
    EXTENDED_CTRLS = {1, 2, 3, 11, 12, 13, 14, 15}
    
    # 확장 컨트롤 ID (역순 저장됨)
    CTRL_ID_TABLE = b' lbt'     # 'tbl ' reversed
    CTRL_ID_FOOTNOTE = b'  nf'  # 'fn  ' reversed
    CTRL_ID_ENDNOTE = b'  ne'   # 'en  ' reversed
    
    @staticmethod
    def normalize_tag(tag_id: int) -> int:
        """Shifted tag를 표준 tag로 변환"""
        # 66-77 범위는 shifted tags
        if 66 <= tag_id <= 77:
            # LIST_HEADER는 예외 (72 -> 24)
            if tag_id == 72:
                return 24
            # 나머지는 -50
            return tag_id - 50
        return tag_id
    
    def __init__(self):
        self.char_shapes: Dict[int, TextStyle] = {}
        self.para_shapes: Dict[int, dict] = {}
        self.font_names: Dict[int, str] = {}
        self._is_compressed: bool = True
        self._footnote_counter: int = 0
        self._base_font_size: int = 1000  # 기본 글자 크기 (10pt)
        self._image_counter: int = 0  # 이미지 삽입 순서 추적
    
    def parse(self, file_path: str) -> Document:
        """HWP 파일 파싱
        
        Args:
            file_path: HWP 파일 경로
            
        Returns:
            Document: 파싱된 문서 객체
        """
        doc = Document()
        
        # 인스턴스 변수 초기화 (재사용 시 이전 결과 제거)
        self.char_shapes.clear()
        self.para_shapes.clear()
        self.font_names.clear()
        self._is_compressed = True
        self._footnote_counter = 0
        self._base_font_size = 1000
        self._image_counter = 0
        
        ole = olefile.OleFileIO(file_path)
        
        try:
            # 0. FileHeader에서 압축 여부 확인
            self._check_file_header(ole)
            
            # 1. DocInfo에서 서식 정보 로드
            self._load_doc_info(ole)
            
            # 2. 기본 글자 크기 결정 (첫 번째 CharShape 기준)
            if self.char_shapes:
                first_shape = self.char_shapes.get(0)
                if first_shape and first_shape.font_size:
                    self._base_font_size = first_shape.font_size
            
            # 3. BinData에서 이미지 먼저 추출 (섹션 파싱 전)
            self._extract_images(ole, doc)
            
            # 4. Section 스트림들 파싱
            entries = ole.listdir()
            section_entries = sorted([
                '/'.join(entry) for entry in entries
                if len(entry) == 2 and entry[0] == 'BodyText' and entry[1].startswith('Section')
            ])
            
            for path in section_entries:
                data = ole.openstream(path).read()
                
                # 압축 해제 시도
                if self._is_compressed:
                    try:
                        data = zlib.decompress(data, -15)  # raw deflate
                    except zlib.error:
                        pass  # 압축 안 된 경우
                
                section = self._parse_section(data, doc)
                doc.sections.append(section)
            
            # 5. 추출된 이미지를 문서 끝에 추가
            # (HWP 배포용 파일은 CTRL_PICTURE_TABLE이 없을 수 있음)
            if doc.sections and doc.images:
                sorted_image_ids = sorted(doc.images.keys())
                for image_id in sorted_image_ids:
                    doc.sections[-1].elements.append(doc.images[image_id])
            
        finally:
            ole.close()
        
        return doc
    
    def _extract_images(self, ole, doc: Document) -> None:
        """BinData 스트림에서 이미지 추출"""
        image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.wmf', '.emf'}
        
        for entry in ole.listdir():
            if len(entry) != 2 or entry[0] != 'BinData':
                continue
                
            file_name = entry[1]
            ext = ('.' + file_name.rsplit('.', 1)[-1].lower()) if '.' in file_name else ''
            
            if ext not in image_extensions:
                continue
            
            try:
                path = '/'.join(entry)
                image_data = ole.openstream(path).read()
                
                # 압축 해제 시도
                if self._is_compressed:
                    try:
                        image_data = zlib.decompress(image_data, -15)
                    except zlib.error:
                        pass
                
                # 이미지 ID
                image_id = file_name.rsplit('.', 1)[0] if '.' in file_name else file_name
                image_format = ext[1:]
                
                # MIME 타입 (Gemini API 미지원 포맷은 분석 스킵)
                mime_map = {
                    'png': 'image/png',
                    'jpg': 'image/jpeg',
                    'jpeg': 'image/jpeg',
                    'gif': 'image/gif',
                    'bmp': 'image/bmp',
                    'tif': 'image/tiff',
                    'tiff': 'image/tiff',
                    'wmf': None,  # Gemini 미지원
                    'emf': None,  # Gemini 미지원
                }
                mime_type = mime_map.get(image_format, 'image/png')
                
                # Gemini Vision API로 이미지 분석 (지원 포맷만)
                description = None
                if mime_type:  # None이면 Gemini 미지원 포맷
                    try:
                        from .. import image_analyzer
                        if image_analyzer.is_available():
                            description = image_analyzer.analyze_image(image_data, mime_type)
                    except Exception:
                        pass
                
                # Image 객체 생성
                image = Image(
                    id=image_id,
                    data=image_data,
                    format=image_format,
                    alt_text=f'Image: {file_name}',
                    description=description
                )
                
                doc.images[image_id] = image
                
            except Exception:
                pass
    
    def _check_file_header(self, ole) -> None:
        """FileHeader에서 압축 여부 확인"""
        if not ole.exists('FileHeader'):
            return
        
        try:
            data = ole.openstream('FileHeader').read()
            if len(data) >= 40:
                # 속성 플래그 (offset 36, DWORD)
                flags = struct.unpack('<I', data[36:40])[0]
                self._is_compressed = bool(flags & 0x01)  # bit 0: 압축 여부
        except Exception:
            pass
    
    def _load_doc_info(self, ole) -> None:
        """DocInfo에서 CharShape, ParaShape, FaceName 로드"""
        if not ole.exists('DocInfo'):
            return
        
        data = ole.openstream('DocInfo').read()
        
        # 압축 해제 시도
        if self._is_compressed:
            try:
                data = zlib.decompress(data, -15)
            except zlib.error:
                pass  # 압축 안 된 경우
        
        pos = 0
        face_name_idx = 0
        char_shape_idx = 0
        para_shape_idx = 0
        
        while pos < len(data) - 4:
            # 레코드 헤더 읽기
            header = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
            
            tag_id = header & 0x3FF
            size = (header >> 20) & 0xFFF
            
            if size == 0xFFF:
                if pos + 4 > len(data):
                    break
                size = struct.unpack('<I', data[pos:pos+4])[0]
                pos += 4
            
            if pos + size > len(data):
                break
            
            record_data = data[pos:pos+size]
            pos += size
            
            # FACE_NAME (19) 파싱
            if tag_id == self.HWPTAG_FACE_NAME and len(record_data) >= 3:
                font_name = self._parse_face_name(record_data)
                if font_name:
                    self.font_names[face_name_idx] = font_name
                face_name_idx += 1
            
            # CHAR_SHAPE (21) 파싱
            elif tag_id == self.HWPTAG_CHAR_SHAPE and len(record_data) >= 72:
                style = self._parse_char_shape(record_data)
                self.char_shapes[char_shape_idx] = style
                char_shape_idx += 1
            
            # PARA_SHAPE (22) 파싱
            elif tag_id == self.HWPTAG_PARA_SHAPE and len(record_data) >= 8:
                para_info = self._parse_para_shape(record_data)
                self.para_shapes[para_shape_idx] = para_info
                para_shape_idx += 1
    
    def _parse_face_name(self, data: bytes) -> Optional[str]:
        """FACE_NAME 레코드에서 글꼴 이름 추출"""
        try:
            # 속성 (1 byte) + 글꼴 이름 (null-terminated UTF-16LE)
            if len(data) < 3:
                return None
            
            # 글꼴 이름 시작 위치
            name_start = 1
            
            # null-terminated UTF-16LE 문자열 읽기
            name_chars = []
            i = name_start
            while i + 1 < len(data):
                code = struct.unpack('<H', data[i:i+2])[0]
                if code == 0:
                    break
                name_chars.append(chr(code))
                i += 2
            
            return ''.join(name_chars) if name_chars else None
        except Exception:
            return None
    
    def _parse_char_shape(self, data: bytes) -> TextStyle:
        """CHAR_SHAPE 레코드 파싱 (HWP 5.0 사양서 기준)"""
        style = TextStyle()
        
        try:
            # 기준 크기 (offset 42-45, HWPUNIT)
            if len(data) >= 46:
                base_size = struct.unpack('<I', data[42:46])[0]
                # 실제 pt = base_size / 100
                style.font_size = base_size / 100
            
            # 속성 (offset 46-49) ← 사양서 정확한 위치!
            if len(data) >= 50:
                attrs = struct.unpack('<I', data[46:50])[0]
                style.italic = bool(attrs & 0x01)      # bit 0
                style.bold = bool(attrs & 0x02)        # bit 1
                style.underline = bool((attrs >> 2) & 0x03)  # bit 2-3
                style.strike = bool((attrs >> 18) & 0x07)    # bit 18-20
            
            # 글꼴 (offset 0-1: 한글 글꼴 ID)
            if len(data) >= 2:
                hangul_font_id = struct.unpack('<H', data[0:2])[0]
                if hangul_font_id in self.font_names:
                    style.font_name = self.font_names[hangul_font_id]
            
            # 글자 색상 (offset 52-55)
            if len(data) >= 56:
                color_val = struct.unpack('<I', data[52:56])[0]
                r = color_val & 0xFF
                g = (color_val >> 8) & 0xFF
                b = (color_val >> 16) & 0xFF
                style.color = f'#{r:02x}{g:02x}{b:02x}'
        except Exception:
            pass
        
        return style
    
    def _parse_para_shape(self, data: bytes) -> dict:
        """PARA_SHAPE 레코드 파싱"""
        info = {'align': 'left', 'outline_level': 0}
        
        try:
            if len(data) >= 4:
                # 속성 플래그
                attrs = struct.unpack('<I', data[0:4])[0]
                
                # 정렬 (bits 0-2)
                align_code = attrs & 0x07
                align_map = {0: 'justify', 1: 'left', 2: 'right', 3: 'center'}
                info['align'] = align_map.get(align_code, 'left')
                
            if len(data) >= 8:
                # 개요 수준 (bits in byte 7 or different location)
                info['outline_level'] = data[7] if data[7] < 10 else 0
        except Exception:
            pass
        
        return info
    
    def _parse_section(self, data: bytes, doc: Document) -> Section:
        """Section 데이터 파싱 (표, 각주 포함)"""
        section = Section()
        records = self._parse_records(data) # records 변수 정의 추가
        # 섹션의 레코드 파싱 (사용자 가이드 기준)
        i = 0
        while i < len(records):
            tag_id, record_data, level = records[i]
            
            # PARA_HEADER: 일반 문단
            if tag_id == self.HWPTAG_PARA_HEADER:
                para_info = self._parse_para_header(record_data)
                text = ''
                char_positions = []
                
                # 다음 레코드들 수집
                base_level = level
                j = i + 1
                while j < len(records):
                    next_tag, next_data, next_level = records[j]
                    
                    # CTRL_HEADER나 다른 PARA_HEADER를 만나면 문단 종료
                    if next_tag == self.HWPTAG_CTRL_HEADER or next_tag == self.HWPTAG_PARA_HEADER:
                        break
                    
                    # level이 base_level 이하로 떨어지고 PARA_TEXT가 아니면 문단 종료
                    if next_level <= base_level and next_tag != self.HWPTAG_PARA_TEXT:
                        break
                    
                    if next_tag == self.HWPTAG_PARA_TEXT:
                        text, ctrl_info = self._extract_text_with_ctrls(next_data, para_info.get('char_count', 0))
                    elif next_tag == self.HWPTAG_PARA_CHAR_SHAPE:
                        char_shape_count = para_info.get('char_shape_count', 0)
                        char_positions = self._parse_para_char_shape(next_data, char_shape_count)
                    
                    j += 1
                
                # 일반 문단 생성 (테이블은 CTRL_HEADER tbl로만 처리)
                if text.strip() and self._is_valid_paragraph(text.strip()):
                    para = self._create_paragraph(text, char_positions, para_info)
                    section.elements.append(para)
                i = j
            
            # CTRL_HEADER: 표, 그림 등 객체
            elif tag_id == self.HWPTAG_CTRL_HEADER:
                if len(record_data) >= 4:
                    ctrl_id = struct.unpack('<I', record_data[0:4])[0]
                    
                    if ctrl_id == 0x74626C20:  # 'tbl ' - 표
                        table, consumed = self._parse_table_from_records_v2(records, i)
                        if table and not table.is_empty():
                            section.elements.append(table)
                        i += consumed
                    
                    elif ctrl_id == 0x24706963:  # '$pic' - 그림
                        picture_info, consumed = self._parse_picture_from_records(records, i)
                        if picture_info:
                            # bin_data_id로 이미지 삽입
                            bin_data_id = picture_info.get('bin_data_id', 0)
                            if bin_data_id > 0:
                                # BIN0001, BIN0002 형식으로 찾기
                                image_key = f'BIN{bin_data_id:04X}'
                                if image_key in doc.images:
                                    section.elements.append(doc.images[image_key])
                        i += consumed
                    
                    else:
                        i += 1
                else:
                    i += 1
            
            else:
                i += 1
        
        return section
    
    def _parse_records(self, data: bytes) -> List[Tuple[int, bytes, int]]:
        """바이트 스트림을 레코드 리스트로 파싱
        
        Returns:
            List of (normalized_tag_id, record_data, level)
        """
        records = []
        pos = 0
        
        while pos < len(data) - 4:
            header = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
            
            original_tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF
            
            # 크기가 0xFFF이면 다음 4바이트가 실제 크기
            if size == 0xFFF:
                if pos + 4 > len(data):
                    break
                size = struct.unpack('<I', data[pos:pos+4])[0]
                pos += 4
            
            if pos + size > len(data):
                break
            
            record_data = data[pos:pos+size]
            pos += size
            
            # Shifted tag를 표준 tag로 변환
            normalized_tag_id = self.normalize_tag(original_tag_id)
            
            records.append((normalized_tag_id, record_data, level))
        
        return records
    
    def _parse_para_header(self, data: bytes) -> dict:
        """PARA_HEADER 레코드 파싱"""
        info = {'char_count': 0, 'para_shape_id': 0}
        
        try:
            if len(data) >= 4:
                char_count = struct.unpack('<I', data[0:4])[0]
                info['char_count'] = char_count & 0x7FFFFFFF
                info['is_last'] = bool(char_count & 0x80000000)
            
            if len(data) >= 10:
                info['para_shape_id'] = struct.unpack('<H', data[8:10])[0]
        except Exception:
            pass
        
        return info
    
    def _parse_para_char_shape(self, data: bytes, count: int = 0) -> List[Tuple[int, int]]:
        """PARA_CHAR_SHAPE 레코드 파싱 (사양서 기준)
        
        Returns:
            List of (position, char_shape_id) tuples
        """
        positions = []
        
        try:
            # count가 0이면 데이터 크기로 계산
            if count == 0:
                count = len(data) // 8
            
            for i in range(count):
                offset = i * 8
                if offset + 8 > len(data):
                    break
                pos = struct.unpack('<I', data[offset:offset+4])[0]
                shape_id = struct.unpack('<I', data[offset+4:offset+8])[0]
                positions.append((pos, shape_id))
        except Exception:
            pass
        
        return positions
    
    def _extract_text_with_ctrls(self, data: bytes, char_count: int = 0) -> Tuple[str, List[Tuple[int, int, bytes]]]:
        """PARA_TEXT 레코드에서 텍스트와 제어문자 정보 추출 (HWP 5.0 사양서 완전 구현)"""
        chars = []
        ctrl_info = []
        pos = 0
        char_pos = 0
        
        while pos < len(data):
            if char_count > 0 and char_pos >= char_count:
                break
            if pos + 2 > len(data):
                break
            
            char_code = struct.unpack('<H', data[pos:pos+2])[0]
            
            if char_code < 32:  # 제어문자
                if char_code in self.CTRL_CHAR_TYPE:
                    # char 타입: 1 WCHAR, 위치 1개
                    if char_code == 10:
                        chars.append('\n')
                    elif char_code == 13:
                        pass
                    elif char_code == 30:
                        chars.append('\u00A0')
                    elif char_code == 31:
                        chars.append(' ')
                    elif char_code == 24:
                        chars.append('-')
                    ctrl_info.append((char_pos, char_code, b''))
                    pos += 2
                    char_pos += 1
                
                elif char_code in self.CTRL_INLINE_TYPE:
                    # inline 타입: 8 WCHAR, 위치 1개
                    ctrl_data = data[pos:pos+16] if pos + 16 <= len(data) else data[pos:]
                    if char_code == 9:
                        chars.append('\t')
                    ctrl_info.append((char_pos, char_code, ctrl_data))
                    pos += 16
                    char_pos += 1
                
                elif char_code in self.CTRL_EXTENDED_TYPE:
                    # extended 타입: 8 WCHAR, 위치 차지 안함!
                    ctrl_data = data[pos:pos+16] if pos + 16 <= len(data) else data[pos:]
                    ctrl_info.append((char_pos, char_code, ctrl_data))
                    pos += 16
                    # char_pos 증가하지 않음 - 중요!
                
                else:
                    pos += 2
                    char_pos += 1
            else:
                chars.append(chr(char_code))
                pos += 2
                char_pos += 1
        
        text = ''.join(chars)
        text = self._clean_text(text)
        return text, ctrl_info
    
    def _clean_text(self, text: str) -> str:
        """HWP에서 추출한 텍스트 정리 - 파싱 오류로 생긴 이상한 문자만 제거"""
        cleaned_chars = []
        for char in text:
            code = ord(char)
            
            # HWP 파싱 오류로 나타나는 특정 CJK 한자들만 제거
            # 潴(U+6F74), 景(U+666F), 慴(U+6174)
            # 주의: 한글 옛 자모(ᄒᆞᆫ글 등)는 원본 문서의 정상 텍스트이므로 유지
            if code in {0x6F74, 0x666F, 0x6174}:
                continue
            
            cleaned_chars.append(char)
        
        return ''.join(cleaned_chars)
    
    def _is_valid_paragraph(self, text: str) -> bool:
        """문단 유효성 검사 - HWP 특수 마커 필터링"""
        # 짧은 문단만 검사
        if len(text) <= 3:
            # HWP 그래픽 요소 마커 (선, 도형 등)
            # 이들은 실제로는 ASCII 코드의 조합(예: 'pn' -> 0x6e70)이 한자로 오인된 것들입니다.
            # 湰(0x6e70), 桤(0x6824), 湯(0x6e6f), 湷(0x6e37), 捤(0x6364), 獥(0x7365) 등
            GRAPHIC_MARKERS = {
                0x6e70, 0x6824, 0x6e6f, 0x6e37, 
                0x6e30, 0xf0e8, 0x6364, 0x7365
            }
            
            for char in text:
                if ord(char) in GRAPHIC_MARKERS:
                    return False
        return True
    
    def _create_paragraph(self, text: str, char_positions: List[Tuple[int, int]], 
                          para_info: dict) -> Paragraph:
        """텍스트와 스타일 정보로 Paragraph 생성"""
        para = Paragraph()
        
        if not char_positions:
            # 스타일 정보 없으면 전체 텍스트를 하나의 run으로
            para.runs.append(TextRun(text=text))
        else:
            # 스타일 적용 위치에 따라 run 분리
            sorted_positions = sorted(char_positions, key=lambda x: x[0])
            
            last_end = 0
            for i, (start_pos, shape_id) in enumerate(sorted_positions):
                # 다음 위치 결정
                if i + 1 < len(sorted_positions):
                    end_pos = sorted_positions[i + 1][0]
                else:
                    end_pos = len(text)
                
                if start_pos < len(text):
                    run_text = text[start_pos:end_pos]
                    if run_text:
                        style = self.char_shapes.get(shape_id, TextStyle())
                        para.runs.append(TextRun(text=run_text, style=style))
            
            # 첫 번째 위치 이전 텍스트
            if sorted_positions and sorted_positions[0][0] > 0:
                first_text = text[:sorted_positions[0][0]]
                if first_text:
                    para.runs.insert(0, TextRun(text=first_text))
        
        # 제목 레벨 감지 (휴리스틱)
        para.heading_level = self._detect_heading_level(para, para_info)
        
        return para
    
    def _collect_list_headers_as_table(self, records: List[Tuple[int, bytes, int]], start_idx: int, base_level: int) -> Tuple[Table, int]:
        """제어문자 11 다음의 LIST_HEADERs를 수집하여 표 생성
        
        이 파일 포맷: PARA_TEXT(ctrl11) → LIST_HEADER들
        """
        table = Table()
        cells = []
        current_cell = None
        consumed = 0
        
        try:
            i = start_idx
            
            while i < len(records):
                tag_id, record_data, level = records[i]
                
                # base_level보다 낮으면 종료 (같으면 계속)
                if level < base_level:
                    break
                
                # 다음 CTRL_HEADER(다른 테이블 시작)를 만나면 종료
                if tag_id == self.HWPTAG_CTRL_HEADER:
                    break
                
                if tag_id == self.HWPTAG_LIST_HEADER:
                    # 새 셀 시작
                    if current_cell is not None:
                        cells.append(current_cell)
                    current_cell = {'paragraphs': []}
                    consumed += 1
                
                elif tag_id == self.HWPTAG_PARA_HEADER:
                    if current_cell is not None:
                        # 셀 내 문단
                        para_info = self._parse_para_header(record_data)
                        text = ''
                        char_positions = []
                        
                        # 다음 PARA_TEXT, PARA_CHAR_SHAPE 찾기
                        j = i + 1
                        while j < len(records):
                            next_tag, next_data, next_level = records[j]
                            
                            if next_tag == self.HWPTAG_PARA_TEXT:
                                text, _ = self._extract_text_with_ctrls(next_data, para_info.get('char_count', 0))
                                consumed += 1
                            elif next_tag == self.HWPTAG_PARA_CHAR_SHAPE:
                                char_shape_count = para_info.get('char_shape_count', 0)
                                char_positions = self._parse_para_char_shape(next_data, char_shape_count)
                                consumed += 1
                            elif next_tag in (self.HWPTAG_PARA_HEADER, self.HWPTAG_LIST_HEADER):
                                break
                            elif next_level < base_level:
                                break
                            else:
                                consumed += 1
                            j += 1
                        
                        if text.strip():
                            para = self._create_paragraph(text, char_positions, para_info)
                            current_cell['paragraphs'].append(para)
                    consumed += 1
                
                else:
                    consumed += 1
                
                i += 1
            
            # 마지막 셀 저장
            if current_cell is not None:
                cells.append(current_cell)
            
            # 셀들을 표로 구성 (3열로 가정)
            if cells:
                table.col_count = 3
                row_count = (len(cells) + 2) // 3
                
                for _ in range(row_count):
                    row = TableRow()
                    for _ in range(table.col_count):
                        row.cells.append(TableCell())
                    table.rows.append(row)
                
                # 셀 배치
                for idx, cell_data in enumerate(cells):
                    if idx >= len(table.rows) * table.col_count:
                        break
                    row_idx = idx // table.col_count
                    col_idx = idx % table.col_count
                    if row_idx < len(table.rows):
                        table.rows[row_idx].cells[col_idx].paragraphs = cell_data['paragraphs']
        
        except Exception:
            pass
        
        return table, consumed
    
    def _detect_heading_level(self, para: Paragraph, para_info: dict) -> HeadingLevel:
        """글자 크기 기반 제목 레벨 감지"""
        # 개요 수준이 있으면 사용
        outline_level = para_info.get('outline_level', 0)
        if 1 <= outline_level <= 6:
            return HeadingLevel(outline_level)
        
        # 글자 크기 기반 휴리스틱
        if para.runs:
            first_run = para.runs[0]
            if first_run.style and first_run.style.font_size:
                font_size = first_run.style.font_size
                
                # 기본 크기 대비 비율로 제목 레벨 결정
                ratio = font_size / self._base_font_size if self._base_font_size else 1.0
                
                if ratio >= 2.0:  # 20pt 이상 (기본 10pt 기준)
                    return HeadingLevel.H1
                elif ratio >= 1.6:  # 16pt 이상
                    return HeadingLevel.H2
        return HeadingLevel.NONE
    
    def _parse_table_from_records_v2(self, records: List[Tuple[int, bytes, int]], start_idx: int) -> Tuple[Table, int]:
        """레코드 리스트에서 표 파싱 (셀 위치/병합 정보 포함)
        
        구조: CTRL_HEADER → TABLE → LIST_HEADERs→ PARAs
        LIST_HEADER에서 row, col, rowspan, colspan 정보 추출
        """
        if start_idx >= len(records):
            return Table(), 0
        
        base_level = records[start_idx][2]  # CTRL_HEADER의 level
        table = Table()
        cells = []  # {'row': int, 'col': int, 'rowspan': int, 'colspan': int, 'paragraphs': []}
        current_cell = None
        consumed = 1  # CTRL_HEADER
        
        try:
            i = start_idx + 1
            table_info = None
            row_count = 0
            col_count = 0
            
            while i < len(records):
                tag_id, record_data, level = records[i]
                
                # level이 base_level 이하면 표 종료
                if level <= base_level:
                    break
                
                # TABLE 레코드에서 row/col 읽기
                if tag_id == self.HWPTAG_TABLE:
                    if len(record_data) >= 8:
                        row_count = struct.unpack('<H', record_data[4:6])[0]
                        col_count = struct.unpack('<H', record_data[6:8])[0]
                        table.col_count = col_count
                        table_info = {'rows': row_count, 'cols': col_count}
                        
                        # 행/셀 초기화
                        for _ in range(row_count):
                            row = TableRow()
                            for _ in range(col_count):
                                row.cells.append(TableCell())
                            table.rows.append(row)
                    consumed += 1
                
                elif tag_id == self.HWPTAG_LIST_HEADER:
                    # 이전 셀 저장
                    if current_cell is not None:
                        cells.append(current_cell)
                    
                    # 셀 개수 확인 - 모든 셀을 읽었으면 종료
                    if table_info and len(cells) >= row_count * col_count:
                        break
                    
                    current_cell = {'paragraphs': []}
                    consumed += 1
                
                elif tag_id == self.HWPTAG_PARA_HEADER:
                    if current_cell is not None:
                        # 셀 내 문단
                        para_info = self._parse_para_header(record_data)
                        text = ''
                        char_positions = []
                        
                        # 다음 PARA_TEXT, PARA_CHAR_SHAPE 찾기
                        j = i + 1
                        while j < len(records):
                            next_tag, next_data, next_level = records[j]
                            
                            if next_tag == self.HWPTAG_PARA_TEXT:
                                text, _ = self._extract_text_with_ctrls(next_data, para_info.get('char_count', 0))
                                consumed += 1
                            elif next_tag == self.HWPTAG_PARA_CHAR_SHAPE:
                                char_shape_count = para_info.get('char_shape_count', 0)
                                char_positions = self._parse_para_char_shape(next_data, char_shape_count)
                                consumed += 1
                            elif next_tag in (self.HWPTAG_PARA_HEADER, self.HWPTAG_LIST_HEADER):
                                break
                            elif next_level <= base_level:
                                break
                            else:
                                consumed += 1
                            j += 1
                        
                        if text.strip():
                            para = self._create_paragraph(text, char_positions, para_info)
                            current_cell['paragraphs'].append(para)
                    consumed += 1
                
                else:
                    consumed += 1
                
                i += 1
            
            # 마지막 셀 저장
            if current_cell is not None:
                cells.append(current_cell)
            
            # 셀을 표에 순차 배치
            if table_info and table.rows:
                for idx, cell_data in enumerate(cells):
                    if idx >= len(table.rows) * table.col_count:
                        break
                    row_idx = idx // table.col_count
                    col_idx = idx % table.col_count
                    if row_idx < len(table.rows):
                        table.rows[row_idx].cells[col_idx].paragraphs = cell_data['paragraphs']
        
        except Exception:
            pass
        
        return table, consumed
    
    def _parse_table_from_ctrl(self, records: List[Tuple[int, bytes, int]], start_idx: int) -> Tuple[Table, int]:
        """CTRL_HEADER부터 시작하는 표 파싱 (사양서 기준)
        
        구조: CTRL_HEADER(ctrl_id='tbl ') → LIST_HEADER들 (각 셀) → PARA_HEADER/PARA_TEXT들
        """
        table = Table()
        consumed = 0
        
        if start_idx >= len(records):
            return table, 0
        
        try:
            # CTRL_HEADER 확인
            tag_id, record_data, level = records[start_idx]
            if tag_id != self.HWPTAG_CTRL_HEADER:
                return table, 0
            
            consumed += 1
            
            # 다음 레코드들에서 LIST_HEADER 수집
            cells = []
            current_cell = None
            idx = start_idx + 1
            
            # 셀 개수를 먼저 세기 (LIST_HEADER 개수)
            list_header_count = 0
            temp_idx = idx
            while temp_idx < len(records):
                temp_tag, _, _ = records[temp_idx]
                if temp_tag == self.HWPTAG_LIST_HEADER:
                    list_header_count += 1
                elif temp_tag == self.HWPTAG_PARA_HEADER and list_header_count > 0:
                    pass  # 셀 내부 문단
                elif temp_tag not in (self.HWPTAG_PARA_TEXT, self.HWPTAG_PARA_CHAR_SHAPE, self.HWPTAG_PARA_LINE_SEG):
                    break
                temp_idx += 1
            
            # 표 구조 추정 (가로 레이아웃 가정)
            if list_header_count > 0:
                # 간단한 휴리스틱: 셀이 4개 미만이면 2열, 그 외는 적절히 분배
                if list_header_count <= 2:
                    table.col_count = list_header_count
                    row_count = 1
                elif list_header_count <= 6:
                    table.col_count = 2
                    row_count = (list_header_count + 1) // 2
                else:
                    table.col_count = 3
                    row_count = (list_header_count + 2) // 3
                
                # 행/셀 초기화
                for _ in range(row_count):
                    row = TableRow()
                    for _ in range(table.col_count):
                        row.cells.append(TableCell())
                    table.rows.append(row)
            
            cell_idx = 0
            
            while idx < len(records) and cell_idx < list_header_count:
                tag_id, record_data, _ = records[idx]
                
                if tag_id == self.HWPTAG_LIST_HEADER:
                    # 새 셀 시작
                    if cell_idx < list_header_count and table.rows:
                        row_idx = cell_idx // table.col_count
                        col_idx = cell_idx % table.col_count
                        if row_idx < len(table.rows):
                            current_cell = table.rows[row_idx].cells[col_idx]
                        cell_idx += 1
                    consumed += 1
                
                elif tag_id == self.HWPTAG_PARA_HEADER and current_cell is not None:
                    # 셀 내 문단
                    para_info = self._parse_para_header(record_data)
                    text = ''
                    char_positions = []
                    
                    # PARA_TEXT 찾기
                    j = idx + 1
                    while j < len(records):
                        next_tag, next_data, _ = records[j]
                        if next_tag == self.HWPTAG_PARA_TEXT:
                            text, _ = self._extract_text_with_ctrls(next_data, para_info.get('char_count', 0))
                            consumed += 1
                        elif next_tag ==self.HWPTAG_PARA_CHAR_SHAPE:
                            char_shape_count = para_info.get('char_shape_count', 0)
                            char_positions = self._parse_para_char_shape(next_data, char_shape_count)
                            consumed += 1
                        elif next_tag in (self.HWPTAG_PARA_HEADER, self.HWPTAG_LIST_HEADER):
                            break
                        else:
                            consumed += 1
                        j += 1
                    
                    if text.strip():
                        para = self._create_paragraph(text, char_positions, para_info)
                        current_cell.paragraphs.append(para)
                    consumed += 1
                
                else:
                    if cell_idx >= list_header_count and tag_id not in (
                        self.HWPTAG_PARA_TEXT, self.HWPTAG_PARA_CHAR_SHAPE, self.HWPTAG_PARA_LINE_SEG
                    ):
                        break
                    consumed += 1
                
                idx += 1
        
        except Exception:
            pass
        
        return table, consumed
    
    def _parse_picture_from_records(self, records: List[Tuple[int, bytes, int]], start_idx: int) -> Tuple[dict, int]:
        """레코드 리스트에서 그림 파싱 (사용자 가이드 기준)
        
        구조: CTRL_HEADER → SHAPE_COMPONENT → SHAPE_COMPONENT_PICTURE
        bin_data_id 추출이 목표
        """
        if start_idx >= len(records):
            return None, 0
        
        base_level = records[start_idx][2]
        consumed = 1  # CTRL_HEADER
        bin_data_id = 0
        
        try:
            i = start_idx + 1
            
            while i < len(records):
                tag_id, record_data, level = records[i]
                
                # level이 base_level 이하면 그림 종료
                if level <= base_level:
                    break
                
                # SHAPE_COMPONENT_PICTURE에서 bin_data_id 추출
                # 사용자 가이드 line 533-616
                if tag_id == 85:  # HWPTAG_SHAPE_COMPONENT_PICTURE의 일반적인 값
                    # bin_data_id는 오프셋 74에 위치 (사용자 가이드 line 581)
                    # 하지만 정확한 파싱은: brightness(1) + contrast(1) + effect(1) + bin_data_id(2)
                    # = 그림 정보 5바이트 중 마지막 2바이트
                    
                    # 간단히: 데이터에서 UINT16 찾기
                    if len(record_data) >= 76:
                        # offset 74에서 2바이트 읽기
                        bin_data_id = struct.unpack('<H', record_data[74:76])[0]
                    consumed += 1
                    break  # bin_data_id 얻었으므로 종료
                
                consumed += 1
                i += 1
        
        except Exception:
            pass
        
        return {'bin_data_id': bin_data_id}, consumed
    
    def _parse_ctrl_header(self, data: bytes) -> dict:
        """CTRL_HEADER 레코드 파싱 (사양서 기준)"""
        if len(data) < 4:
            return None
        
        ctrl_id = struct.unpack('<I', data[0:4])[0]
        
        # ctrl_id를 4글자로 변환
        ctrl_ch = ''.join([
            chr((ctrl_id >> 24) & 0xFF),
            chr((ctrl_id >> 16) & 0xFF),
            chr((ctrl_id >> 8) & 0xFF),
            chr(ctrl_id & 0xFF)
        ])
        
        return {
            'ctrl_id': ctrl_id,
            'ctrl_ch': ctrl_ch,  # 'tbl ', '$pic', 'secd' 등
            'data': data[4:]
        }
    
    def _parse_table_contents(self, data: bytes, records: List[Tuple[int, bytes, int]], start_idx: int) -> Tuple[Table, int]:
        """TABLE 레코드와 후속 셀 데이터로 표 파싱 (사양서 완전 구현)"""
        table = Table()
        consumed = 0
        
        try:
            # TABLE 레코드 파싱 (사양서 463-493줄)
            if len(data) < 8:
                return table, 0
            
            attr = struct.unpack('<I', data[0:4])[0]
            row_count = struct.unpack('<H', data[4:6])[0]
            table.col_count = struct.unpack('<H', data[6:8])[0]
            
            # 행/셀 초기화
            total_cells = row_count * table.col_count
            for _ in range(row_count):
                row = TableRow()
                for _ in range(table.col_count):
                    row.cells.append(TableCell())
                table.rows.append(row)
            
            # LIST_HEADER 및 PARA 수집
            idx = start_idx
            cell_idx = 0
            current_cell = None
            
            while idx < len(records) and cell_idx < total_cells:
                tag_id, record_data, level = records[idx]
                
                if tag_id == self.HWPTAG_LIST_HEADER:
                    # 새 셀 시작
                    if cell_idx < total_cells:
                        curr_row = cell_idx // table.col_count
                        curr_col = cell_idx % table.col_count
                        current_cell = table.rows[curr_row].cells[curr_col]
                        cell_idx += 1
                    consumed += 1
                
                elif tag_id == self.HWPTAG_PARA_HEADER and current_cell is not None:
                    # 셀 내 문단
                    para_info = self._parse_para_header(record_data)
                    text = ''
                    char_positions = []
                    
                    # 다음 PARA_TEXT, PARA_CHAR_SHAPE 찾기
                    j = idx + 1
                    while j < len(records):
                        next_tag, next_data, _ = records[j]
                        if next_tag == self.HWPTAG_PARA_TEXT:
                            text, _ = self._extract_text_with_ctrls(next_data, para_info.get('char_count', 0))
                            consumed += 1
                        elif next_tag == self.HWPTAG_PARA_CHAR_SHAPE:
                            char_shape_count = para_info.get('char_shape_count', 0)
                            char_positions = self._parse_para_char_shape(next_data, char_shape_count)
                            consumed += 1
                        elif next_tag in (self.HWPTAG_PARA_HEADER, self.HWPTAG_LIST_HEADER):
                            break
                        else:
                            consumed += 1
                        j += 1
                    
                    if text.strip():
                        para = self._create_paragraph(text, char_positions, para_info)
                        current_cell.paragraphs.append(para)
                    consumed += 1
                
                else:
                    if cell_idx >= total_cells and tag_id not in (
                        self.HWPTAG_PARA_TEXT, self.HWPTAG_PARA_CHAR_SHAPE, self.HWPTAG_PARA_LINE_SEG
                    ):
                        break
                    consumed += 1
                
                idx += 1
        
        except Exception:
            pass
        
        return table, consumed
    
    def _extract_text(self, data: bytes) -> str:
        """PARA_TEXT 레코드에서 텍스트 추출 (하위 호환)"""
        text, _ = self._extract_text_with_ctrls(data)
        return text
    
    @staticmethod
    def quick_extract(file_path: str) -> str:
        """PrvText에서 빠른 텍스트 추출
        
        Args:
            file_path: HWP 파일 경로
            
        Returns:
            str: 추출된 텍스트
        """
        try:
            ole = olefile.OleFileIO(file_path)
            try:
                if ole.exists('PrvText'):
                    data = ole.openstream('PrvText').read()
                    return data.decode('utf-16-le', errors='ignore')
                return ''
            finally:
                ole.close()
        except Exception:
            return ''
