"""
HWPX 파서

HWPX (ZIP+XML) 형식 파일을 파싱하는 모듈
"""

import io
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Set

from .base import BaseParser
from ..models import (
    Document, Section, Paragraph, TextRun, Table, TableRow, TableCell,
    TextStyle, Footnote, HeadingLevel, Image
)


class HwpxParser(BaseParser):
    """HWPX (ZIP+XML) 파일 파서"""
    
    SUPPORTED_EXTENSIONS: Set[str] = {'.hwpx'}
    
    # XML 네임스페이스
    NS = {
        'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
        'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
        'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
        'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
        'hm': 'http://www.hancom.co.kr/hwpml/2011/master',
        'hv': 'http://www.hancom.co.kr/hwpml/2011/version',
    }
    
    def __init__(self):
        self.char_shapes: Dict[str, TextStyle] = {}  # id -> TextStyle
        self.para_shapes: Dict[str, dict] = {}       # id -> {align, ...}
        self.font_faces: Dict[str, str] = {}         # id -> font name
        self._footnote_counter: int = 0
        self._base_font_size: int = 1000  # 기본 글자 크기 (10pt = 1000)
        self._analyze_images: bool = True  # 이미지 분석 여부

    def parse(self, file_path: str, analyze_images: bool = True) -> Document:
        """HWPX 파일 파싱

        Args:
            file_path: HWPX 파일 경로
            analyze_images: Gemini API로 이미지 분석 수행 여부

        Returns:
            Document: 파싱된 문서 객체
        """
        doc = Document()
        self._analyze_images = analyze_images

        # 인스턴스 변수 초기화 (재사용 시 이전 결과 제거)
        self.char_shapes.clear()
        self.para_shapes.clear()
        self.font_faces.clear()
        self._footnote_counter = 0
        self._base_font_size = 1000
        
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 1. 네임스페이스 추출 (동적)
            self._extract_namespaces(zf)
            
            # 2. header.xml에서 서식 정보 로드
            self._load_header(zf)
            
            # 3. 기본 글자 크기 결정 (첫 번째 CharShape 기준)
            if self.char_shapes and '0' in self.char_shapes:
                first_shape = self.char_shapes['0']
                if first_shape.font_size:
                    self._base_font_size = first_shape.font_size
            
            # 4. 이미지 추출 (BinData 폴더)
            self._extract_images(zf, doc)
            
            # 5. 구역 파일 목록
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])
            
            # 6. 각 section 파싱
            for sf in section_files:
                section = self._parse_section(zf, sf, doc)
                doc.sections.append(section)
        
        return doc
    
    def _extract_namespaces(self, zf: zipfile.ZipFile) -> None:
        """header.xml에서 실제 네임스페이스 추출"""
        if 'Contents/header.xml' not in zf.namelist():
            return
        
        xml_data = zf.read('Contents/header.xml')
        for event, elem in ET.iterparse(io.BytesIO(xml_data), events=['start-ns']):
            prefix, uri = elem
            if prefix:
                self.NS[prefix] = uri
    
    def _extract_images(self, zf: zipfile.ZipFile, doc: Document) -> None:
        """BinData 폴더에서 이미지 추출"""
        # 지원하는 이미지 확장자
        image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.wmf', '.emf'}
        
        for file_path in zf.namelist():
            # BinData 폴더 내 파일만 처리
            if not file_path.startswith('BinData/'):
                continue
            
            # 파일 확장자 확인
            ext = '.' + file_path.rsplit('.', 1)[-1].lower() if '.' in file_path else ''
            if ext not in image_extensions:
                continue
            
            try:
                # 이미지 데이터 읽기
                image_data = zf.read(file_path)
                
                # 이미지 ID (파일명에서 추출)
                file_name = file_path.rsplit('/', 1)[-1]
                image_id = file_name.rsplit('.', 1)[0] if '.' in file_name else file_name
                
                # 이미지 형식
                image_format = ext[1:]  # 점 제거
                
                # MIME 타입 결정 (Gemini API 미지원 포맷은 분석 스킵)
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
                
                # Gemini Vision API로 이미지 분석 (옵션이 활성화되고 지원 포맷인 경우만)
                description = None
                if self._analyze_images and mime_type:  # 분석 옵션 확인 + None이면 Gemini 미지원 포맷
                    try:
                        from .. import image_analyzer
                        if image_analyzer.is_available():
                            description = image_analyzer.analyze_image(image_data, mime_type)
                    except Exception as e:
                        pass  # 분석 실패 시 무시
                
                # Image 객체 생성
                image = Image(
                    id=image_id,
                    data=image_data,
                    format=image_format,
                    alt_text=f'Image: {file_name}',
                    description=description  # AI 분석 설명
                )
                
                # Document에 추가
                doc.images[image_id] = image
                
            except Exception:
                # 이미지 읽기 실패 시 무시
                pass
    
    def _load_header(self, zf: zipfile.ZipFile) -> None:
        """header.xml에서 charPr, paraPr, fontface 로드"""
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
        
        # 글꼴 정보 (fontface)
        self._load_font_faces(ref_list)
        
        # 글자 모양 (charPr)
        self._load_char_properties(ref_list)
        
        # 문단 모양 (paraPr)
        self._load_para_properties(ref_list)
    
    def _load_font_faces(self, ref_list) -> None:
        """글꼴 정보 로드"""
        fontfaces = ref_list.find('.//hh:fontfaces', self.NS)
        if fontfaces is None:
            fontfaces = ref_list.find('.//{*}fontfaces')
        
        if fontfaces is None:
            return
        
        for fontface in fontfaces.findall('.//{*}font'):
            font_id = fontface.get('id', '')
            font_name = fontface.get('face', '')
            if font_id and font_name:
                self.font_faces[font_id] = font_name
    
    def _load_char_properties(self, ref_list) -> None:
        """글자 모양 정보 로드"""
        char_props = ref_list.find('.//hh:charProperties', self.NS)
        if char_props is None:
            char_props = ref_list.find('.//{*}charProperties')
        
        if char_props is None:
            return
        
        for cp in char_props.findall('.//{*}charPr'):
            cp_id = cp.get('id', '0')
            style = TextStyle()
            
            # 속성 파싱
            style.font_size = int(cp.get('height', '1000'))
            style.color = cp.get('textColor', '#000000')
            style.bold = cp.get('bold', 'false').lower() == 'true'
            style.italic = cp.get('italic', 'false').lower() == 'true'
            style.underline = cp.get('underline', 'false').lower() == 'true'
            style.strike = cp.get('strikeout', 'false').lower() == 'true'
            
            # 글꼴 참조
            font_ref = cp.find('.//{*}fontRef')
            if font_ref is not None:
                hangul_font_id = font_ref.get('hangul', '')
                if hangul_font_id in self.font_faces:
                    style.font_name = self.font_faces[hangul_font_id]
            
            self.char_shapes[cp_id] = style
    
    def _load_para_properties(self, ref_list) -> None:
        """문단 모양 정보 로드"""
        para_props = ref_list.find('.//hh:paraProperties', self.NS)
        if para_props is None:
            para_props = ref_list.find('.//{*}paraProperties')
        
        if para_props is None:
            return
        
        for pp in para_props.findall('.//{*}paraPr'):
            pp_id = pp.get('id', '0')
            self.para_shapes[pp_id] = {
                'align': pp.get('align', 'JUSTIFY'),
                'heading': pp.get('heading', None),
            }
    
    def _get_image_id_from_pic(self, pic_elem) -> str:
        """pic 요소에서 이미지 ID 추출"""
        # imageRect > imgDat 또는 img 요소에서 binItemIdRef 찾기
        for child in pic_elem.iter():
            local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local_name in ('imgDat', 'img', 'imageRect'):
                bin_ref = child.get('binItemIdRef', child.get('binaryItemIDRef', ''))
                if bin_ref:
                    return bin_ref
        return None
    
    def _parse_section(self, zf: zipfile.ZipFile, path: str, doc: Document) -> Section:
        """section*.xml 파싱"""
        section = Section()
        xml_data = zf.read(path)
        root = ET.fromstring(xml_data)
        
        # 이미 처리된 요소 추적
        processed = set()
        
        # 모든 p 요소 순회
        for p_elem in root.iter():
            local_name = p_elem.tag.split('}')[-1] if '}' in p_elem.tag else p_elem.tag
            
            if local_name == 'p' and id(p_elem) not in processed:
                processed.add(id(p_elem))
                
                # 표가 포함된 문단인지 확인
                tbl = self._find_child(p_elem, 'tbl')
                if tbl is not None:
                    table = self._parse_table(tbl, processed)
                    if not table.is_empty():
                        section.elements.append(table)
                else:
                    # 각주 처리
                    self._extract_footnotes(p_elem, doc)
                    
                    # 이미지 참조 확인 (pic 요소)
                    pic = self._find_child(p_elem, 'pic')
                    if pic is not None:
                        # 이미지 ID 추출
                        img_id = self._get_image_id_from_pic(pic)
                        if img_id and img_id in doc.images:
                            section.elements.append(doc.images[img_id])
                    
                    para = self._parse_paragraph(p_elem)
                    # HWP 특수 마커 필터링
                    if not para.is_empty() and self._is_valid_paragraph(para):
                        section.elements.append(para)
        
        return section
    
    def _is_valid_paragraph(self, para: Paragraph) -> bool:
        """문단 유효성 검사 - HWP 특수 마커 필터링"""
        text = ''.join(run.text for run in para.runs).strip()
        if len(text) <= 3:
            # HWP 그래픽 요소 마커 (선, 도형 등)
            # 이들은 실제로는 ASCII 코드의 조합(예: 'pn' -> 0x6e70)이 한자로 오인된 것들입니다.
            GRAPHIC_MARKERS = {
                0x6e70, 0x6824, 0x6e6f, 0x6e37, 
                0x6e30, 0xf0e8, 0x6364, 0x7365
            }
            
            for char in text:
                if ord(char) in GRAPHIC_MARKERS:
                    return False
        return True
    
    def _parse_paragraph(self, p_elem) -> Paragraph:
        """p 요소 파싱"""
        para = Paragraph()
        
        # 문단 스타일 ID로 제목 레벨 추정
        style_id = p_elem.get('styleIDRef', '')
        para_pr_id = p_elem.get('paraPrIDRef', '')
        
        # 첫 번째 run의 charPrIDRef 가져오기
        runs = self._find_all_children(p_elem, 'run')
        first_char_pr_id = runs[0].get('charPrIDRef', '') if runs else ''
        
        for run in runs:
            char_pr_id = run.get('charPrIDRef', '0')
            style = self.char_shapes.get(char_pr_id, TextStyle())
            
            for child in run:
                local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                if local_name == 't':
                    text = self._extract_text(child)
                    if text:
                        # 스타일 복사하여 사용
                        run_style = TextStyle(
                            bold=style.bold,
                            italic=style.italic,
                            underline=style.underline,
                            strike=style.strike,
                            font_size=style.font_size,
                            font_name=style.font_name,
                            color=style.color,
                        )
                        para.runs.append(TextRun(text=text, style=run_style))
        
        # 제목 레벨 감지 (휴리스틱) - runs 파싱 후 실행
        para.heading_level = self._detect_heading_level(style_id, para_pr_id, first_char_pr_id)
        
        return para
    
    def _detect_heading_level(self, style_id: str, para_pr_id: str, 
                                 first_char_pr_id: str = '') -> HeadingLevel:
        """제목 레벨 감지 (휴리스틱)"""
        # styleIDRef 기반 감지
        if style_id:
            style_lower = style_id.lower()
            for level in range(1, 7):
                if f'heading{level}' in style_lower or f'제목{level}' in style_lower:
                    return HeadingLevel(level)
                if f'h{level}' in style_lower:
                    return HeadingLevel(level)
        
        # paraPr의 heading 속성 확인
        if para_pr_id in self.para_shapes:
            heading = self.para_shapes[para_pr_id].get('heading')
            if heading:
                try:
                    level = int(heading)
                    if 1 <= level <= 6:
                        return HeadingLevel(level)
                except (ValueError, TypeError):
                    pass
        
        # 글자 크기 기반 휴리스틱
        if first_char_pr_id and first_char_pr_id in self.char_shapes:
            style = self.char_shapes[first_char_pr_id]
            if style.font_size and self._base_font_size:
                ratio = style.font_size / self._base_font_size
                
                if ratio >= 2.0:  # 20pt 이상 (기본 10pt 기준)
                    return HeadingLevel.H1
                elif ratio >= 1.6:  # 16pt 이상
                    return HeadingLevel.H2
                elif ratio >= 1.4:  # 14pt 이상
                    return HeadingLevel.H3
                elif ratio >= 1.2 and style.bold:  # 12pt + 굵게
                    return HeadingLevel.H4
        
        return HeadingLevel.NONE
    
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
        
        text = ''.join(parts)
        
        # HWP 특수 마커 및 그래픽 앵커 필터링
        # 湰(0x6e30), 桤(0x6824), 湯(0x6e6f), 湷(0x6e37), (0xf0e8)
        # 이들은 텍스트가 아니라 선, 도형 등의 앵커입니다.
        GRAPHIC_MARKERS = {
            0x6e70, 0x6824, 0x6e6f, 0x6e37, 
            0x6e30, 0xf0e8, 0x6364, 0x7365
        }
        
        filtered_chars = []
        for char in text:
            code = ord(char)
            if code in GRAPHIC_MARKERS:
                continue
            # PUA 영역의 다른 마커들도 안전하게 제거 (사용자 정의 영역 제외)
            if 0xe000 <= code <= 0xf8ff:
                continue
                
            filtered_chars.append(char)
            
        return ''.join(filtered_chars)
    
    def _extract_footnotes(self, p_elem, doc: Document) -> None:
        """문단에서 각주 추출"""
        for run in self._find_all_children(p_elem, 'run'):
            for child in run:
                local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                if local_name == 'footNote':
                    self._footnote_counter += 1
                    fn_id = child.get('id', f'fn{self._footnote_counter}')
                    
                    footnote = Footnote(
                        id=fn_id,
                        number=self._footnote_counter,
                        content=[]
                    )
                    
                    # 각주 내 문단들
                    for fn_p in self._find_all_children(child, 'p'):
                        para = self._parse_paragraph(fn_p)
                        footnote.content.append(para)
                    
                    doc.footnotes[fn_id] = footnote
                
                elif local_name == 'endNote':
                    self._footnote_counter += 1
                    en_id = child.get('id', f'en{self._footnote_counter}')
                    
                    endnote = Footnote(
                        id=en_id,
                        number=self._footnote_counter,
                        content=[]
                    )
                    
                    for en_p in self._find_all_children(child, 'p'):
                        para = self._parse_paragraph(en_p)
                        endnote.content.append(para)
                    
                    doc.endnotes[en_id] = endnote
    
    def _parse_table(self, tbl_elem, processed: set) -> Table:
        """tbl 요소 파싱"""
        table = Table()
        table.col_count = int(tbl_elem.get('colCnt', '0'))

        for tr in self._find_all_children(tbl_elem, 'tr'):
            row = TableRow()

            for tc in self._find_all_children(tr, 'tc'):
                cell = TableCell()
                cell.colspan = int(tc.get('colSpan', '1'))
                cell.rowspan = int(tc.get('rowSpan', '1'))

                # 셀 내 문단들 (재귀 탐색)
                # tc 하위의 모든 p 요소를 찾음 (중첩 포함)
                for child in tc.iter():
                    child_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if child_local == 'p':
                        processed.add(id(child))
                        para = self._parse_paragraph(child)
                        cell.paragraphs.append(para)
                    # 셀 내 이미지 추출 (hp:pic > hc:img)
                    elif child_local == 'img':
                        img_ref = child.get('binaryItemIDRef')
                        if img_ref:
                            cell.image_ids.append(img_ref)

                row.cells.append(cell)

            table.rows.append(row)

        return table
    
    def _find_child(self, elem, local_name: str):
        """로컬 이름으로 자식 요소 찾기 (직접 자식만)"""
        for child in elem:
            child_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if child_name == local_name:
                return child
        
        # 한 단계 더 깊이 탐색 (run 내부)
        for child in elem:
            for grandchild in child:
                grandchild_name = grandchild.tag.split('}')[-1] if '}' in grandchild.tag else grandchild.tag
                if grandchild_name == local_name:
                    return grandchild
        
        return None
    
    def _find_all_children(self, elem, local_name: str) -> List:
        """로컬 이름으로 모든 자식 요소 찾기 (직접 자식만)"""
        result = []
        for child in elem:
            child_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if child_name == local_name:
                result.append(child)
        return result
    
    @staticmethod
    def quick_extract(file_path: str) -> str:
        """PrvText.txt에서 빠른 텍스트 추출
        
        Args:
            file_path: HWPX 파일 경로
            
        Returns:
            str: 추출된 텍스트
        """
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                if 'Preview/PrvText.txt' in zf.namelist():
                    return zf.read('Preview/PrvText.txt').decode('utf-8', errors='ignore')
        except Exception:
            pass
        return ''
