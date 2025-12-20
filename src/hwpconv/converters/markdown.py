"""
Markdown 변환기

Document 객체를 Markdown 형식으로 변환
"""

from typing import List, Optional
from .base import BaseConverter
from ..models import Document, Section, Paragraph, Table, TextRun, HeadingLevel, Image


class MarkdownConverter(BaseConverter):
    """Markdown 변환기"""
    
    def __init__(self, include_footnotes: bool = True, 
                 include_metadata: bool = False,
                 include_images: bool = True,
                 heading_style: str = 'atx'):
        """
        Args:
            include_footnotes: 각주를 포함할지 여부
            include_metadata: YAML front matter로 메타데이터 포함 여부
            include_images: 이미지를 포함할지 여부 (Base64 인라인)
            heading_style: 제목 스타일 ('atx' = #, 'setext' = underline)
        """
        self.include_footnotes = include_footnotes
        self.include_metadata = include_metadata
        self.include_images = include_images
        self.heading_style = heading_style
        self._footnote_refs: List[int] = []  # 본문에서 참조된 각주 번호
    
    def convert(self, doc: Document) -> str:
        """Document를 Markdown으로 변환
        
        Args:
            doc: 변환할 Document 객체
            
        Returns:
            str: Markdown 문자열
        """
        lines = []
        self._footnote_refs = []
        
        if self.include_metadata and doc.metadata:
            lines.append('---')
            for key, value in doc.metadata.items():
                # YAML 이스케이프 (백슬래시, 따옴표, 콜론, 줄바꿈 처리)
                value = value.replace('\\', '\\\\')  # 백슬래시 먼저
                if '"' in value:
                    value = value.replace('"', '\\"')
                if ':' in value or '\n' in value or '"' in value:
                    value = f'"{value}"'
                lines.append(f'{key}: {value}')
            lines.append('---')
            lines.append('')
        
        # 각 섹션 변환
        for section in doc.sections:
            section_lines = self._convert_section(section, doc)
            lines.extend(section_lines)
        
        # 각주 추가
        if self.include_footnotes and doc.footnotes:
            footnote_lines = self._convert_footnotes(doc)
            if footnote_lines:
                lines.extend(['', '---', ''])
                lines.extend(footnote_lines)
        
        # 미주 추가
        if self.include_footnotes and doc.endnotes:
            lines.extend(['', '---', '', '## Notes', ''])
            for en_id, en in sorted(doc.endnotes.items(), key=lambda x: x[1].number):
                en_text = en.text.replace('\n', ' ').strip()
                lines.append(f'{en.number}. {en_text}')
        
        # 이미지는 본문에서 올바른 위치에 표시됨 (section.elements에 Image 객체 포함)
        
        # 마지막 빈 줄 정리
        result = '\n'.join(lines)
        while result.endswith('\n\n\n'):
            result = result[:-1]
        
        return result
    
    def _convert_section(self, section: Section, doc: Document) -> List[str]:
        """섹션 변환"""
        lines = []
        prev_was_heading = False
        
        for elem in section.elements:
            if isinstance(elem, Paragraph):
                line = self._convert_paragraph(elem)
                if line:
                    # 제목 뒤에는 빈 줄 추가
                    # 앞쪽 빈 줄 처리 (헤딩 전 등)
                    if prev_was_heading and (not lines or lines[-1] != ''):
                        lines.append('')
                    
                    lines.append(line)
                    
                    # 뒤쪽 빈 줄 (문단 구분)
                    if not lines or lines[-1] != '':
                        lines.append('')
                    
                    prev_was_heading = elem.heading_level != HeadingLevel.NONE
                    
            elif isinstance(elem, Table):
                table_md = self._convert_table(elem)
                if table_md:
                    lines.append(table_md)
                    lines.append('')
                prev_was_heading = False
            
            elif isinstance(elem, Image):
                # 인라인 이미지 (섹션 내)
                if self.include_images:
                    img_md = self._convert_image(elem)
                    if img_md:
                        lines.append(img_md)
                        lines.append('')
                prev_was_heading = False
        
        return lines
    
    def _convert_image(self, img: Image) -> str:
        """이미지 → Markdown (설명만 표시, Base64 제거)"""
        # AI 분석 설명이 있으면 해당 설명만 표시
        if img.description:
            return f'\n> 🖼️ **[이미지]**: {img.description}\n'
        else:
            # 설명이 없으면 이미지 존재만 표시
            return f'\n> 🖼️ **[이미지]**: *(이미지 분석 불가)*\n'
    
    def _convert_paragraph(self, para: Paragraph) -> str:
        """문단 → Markdown"""
        text = ''
        
        for run in para.runs:
            run_text = self._escape_markdown_special(run.text)
            
            # 스타일 적용 (빈 텍스트에는 적용하지 않음)
            if run_text.strip():
                # 볼드+이탤릭 조합
                if run.style.bold and run.style.italic:
                    run_text = self._wrap_style(run_text, '***')
                elif run.style.bold:
                    run_text = self._wrap_style(run_text, '**')
                elif run.style.italic:
                    run_text = self._wrap_style(run_text, '*')
                # 밑줄은 기본적으로 무시 (마크다운에 없음, HTML 사용 시 복잡해짐)
                # underline이 유일한 스타일이면 이탤릭으로 대체
                elif run.style.underline:
                    run_text = self._wrap_style(run_text, '*')
                
                # 취소선은 볼드/이탤릭이 없을 때만 적용
                if run.style.strike and not (run.style.bold or run.style.italic):
                    run_text = self._wrap_style(run_text, '~~')
            
            text += run_text
        
        # 제목 레벨
        if para.heading_level != HeadingLevel.NONE:
            text = text.strip()
            if self.heading_style == 'setext' and para.heading_level.value <= 2:
                # Setext 스타일 (H1은 =, H2는 -)
                underline = '=' if para.heading_level == HeadingLevel.H1 else '-'
                return f'{text}\n{underline * len(text)}'
            else:
                # ATX 스타일 (#)
                prefix = '#' * para.heading_level.value
                return f'{prefix} {text}'
        
        return text
    
    def _wrap_style(self, text: str, marker: str) -> str:
        """스타일 마커로 텍스트 감싸기 (공백 처리)"""
        # 앞뒤 공백 보존
        leading_space = ''
        trailing_space = ''
        
        if text.startswith(' '):
            leading_space = ' '
            text = text[1:]
        if text.endswith(' '):
            trailing_space = ' '
            text = text[:-1]
        
        if text:
            return f'{leading_space}{marker}{text}{marker}{trailing_space}'
        return leading_space + trailing_space
    
    def _escape_markdown_special(self, text: str) -> str:
        """Markdown 특수문자 이스케이프 (최소한만)"""
        # 표 구분자만 이스케이프 (다른 마크업은 의도적일 수 있음)
        return text
    
    def _convert_table(self, table: Table) -> str:
        """표 → Markdown"""
        if not table.rows:
            return ''
        
        lines = []
        
        # 컬럼 수 결정
        col_count = table.col_count
        if col_count == 0:
            col_count = max(len(row.cells) for row in table.rows) if table.rows else 0
        
        if col_count == 0:
            return ''
        
        # 각 행 변환
        for i, row in enumerate(table.rows):
            cells = []
            
            for cell in row.cells:
                # 셀 텍스트 정리
                cell_text = cell.text
                # 줄바꿈을 <br>로 변환 또는 공백으로
                cell_text = cell_text.replace('\n', ' ').strip()
                # 파이프 이스케이프
                cell_text = cell_text.replace('|', '\\|')
                # 빈 셀은 공백으로 (마크다운 렌더러 호환성)
                cells.append(cell_text if cell_text else ' ')
            
            # 컬럼 수 맞추기
            while len(cells) < col_count:
                cells.append('')
            
            # colspan 처리 (병합된 셀은 내용 후 빈 셀 추가)
            # 참고: 기본 Markdown은 colspan을 지원하지 않음
            
            lines.append('| ' + ' | '.join(cells) + ' |')
            
            # 헤더 구분선 (첫 행 다음)
            if i == 0:
                # 정렬 정보가 있으면 적용 (현재는 기본 왼쪽 정렬)
                separators = ['---'] * col_count
                lines.append('| ' + ' | '.join(separators) + ' |')
        
        return '\n'.join(lines)
    
    def _convert_footnotes(self, doc: Document) -> List[str]:
        """각주를 Markdown 형식으로 변환"""
        lines = []
        
        for fn_id, fn in sorted(doc.footnotes.items(), key=lambda x: x[1].number):
            fn_text = fn.text.replace('\n', ' ').strip()
            lines.append(f'[^{fn.number}]: {fn_text}')
        
        return lines
