"""
HTML 변환기
"""

from .base import BaseConverter
from ..models import Document, Paragraph, Table, HeadingLevel, Image


class HtmlConverter(BaseConverter):
    """HTML 변환기"""
    
    def __init__(self, include_images: bool = True, include_footnotes: bool = True):
        """
        Args:
            include_images: 이미지를 포함할지 여부
            include_footnotes: 각주를 포함할지 여부
        """
        self.include_images = include_images
        self.include_footnotes = include_footnotes
    
    def convert(self, doc: Document) -> str:
        parts = ['<!DOCTYPE html>', '<html lang="ko">', '<head>',
                 '<meta charset="UTF-8">', '<title>Document</title>',
                 '<style>',
                 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;',
                 'max-width:800px;margin:0 auto;padding:20px;line-height:1.6;}',
                 'table{border-collapse:collapse;width:100%;margin:1em 0;}',
                 'th,td{border:1px solid #ddd;padding:8px;text-align:left;}',
                 'th{background:#f5f5f5;}',
                 'img{max-width:100%;height:auto;margin:1em 0;border-radius:4px;}',
                 '.image-container{margin:1.5em 0;padding:1em;background:#f9f9f9;border-radius:8px;}',
                 '.image-description{font-size:0.9em;color:#666;margin-top:0.5em;',
                 'padding:0.5em;background:#fff;border-left:3px solid #8B5CF6;}',
                 'blockquote{margin:1em 0;padding:1em;background:#f5f5f5;border-left:4px solid #8B5CF6;}',
                 '</style></head>', '<body>']
        
        for section in doc.sections:
            for elem in section.elements:
                if isinstance(elem, Paragraph):
                    parts.append(self._convert_paragraph(elem))
                elif isinstance(elem, Table):
                    parts.append(self._convert_table(elem))
                elif isinstance(elem, Image) and self.include_images:
                    parts.append(self._convert_image(elem))
        
        # 이미지는 섹션 내에서 이미 표시되므로 중복 표시하지 않음
        # (기존 코드: 문서 끝에 모든 이미지를 다시 표시했음)
        
        # 각주 섹션
        if self.include_footnotes and doc.footnotes:
            parts.append('<hr>')
            parts.append('<h2>Footnotes</h2>')
            parts.append('<ol>')
            for fn_id, fn in sorted(doc.footnotes.items(), key=lambda x: x[1].number):
                fn_text = self._escape_html(fn.text)
                parts.append(f'<li id="fn-{fn.number}">{fn_text}</li>')
            parts.append('</ol>')
        
        parts.extend(['</body>', '</html>'])
        return '\n'.join(parts)
    
    def _convert_paragraph(self, para: Paragraph) -> str:
        text = ''
        for run in para.runs:
            run_text = self._escape_html(run.text)
            if run.style.bold:
                run_text = f'<strong>{run_text}</strong>'
            if run.style.italic:
                run_text = f'<em>{run_text}</em>'
            if run.style.underline:
                run_text = f'<u>{run_text}</u>'
            if run.style.strike:
                run_text = f'<del>{run_text}</del>'
            text += run_text
        
        if para.heading_level != HeadingLevel.NONE:
            lvl = para.heading_level.value
            return f'<h{lvl}>{text}</h{lvl}>'
        return f'<p>{text}</p>'
    
    def _convert_table(self, table: Table) -> str:
        lines = ['<table>']
        for i, row in enumerate(table.rows):
            lines.append('<tr>')
            tag = 'th' if i == 0 else 'td'
            for cell in row.cells:
                attrs = ''
                if cell.colspan > 1:
                    attrs += f' colspan="{cell.colspan}"'
                if cell.rowspan > 1:
                    attrs += f' rowspan="{cell.rowspan}"'
                lines.append(f'<{tag}{attrs}>{self._escape_html(cell.text)}</{tag}>')
            lines.append('</tr>')
        lines.append('</table>')
        return '\n'.join(lines)
    
    def _convert_image(self, img: Image) -> str:
        """이미지 → HTML"""
        alt = self._escape_html(img.alt_text or f'Image {img.id}')
        
        html = '<div class="image-container">'
        html += f'<img src="{img.data_uri}" alt="{alt}">'
        
        # AI 분석 설명이 있으면 포함
        if img.description:
            desc = self._escape_html(img.description)
            html += f'<div class="image-description">{desc}</div>'
        elif img.alt_text:
            html += f'<div class="image-description">{alt}</div>'
        
        html += '</div>'
        return html
    
    def _escape_html(self, text: str) -> str:
        """HTML 특수문자 이스케이프 (XSS 방지)"""
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#x27;'))
