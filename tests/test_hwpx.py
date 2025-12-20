"""
HWPX 파서 테스트
"""

import pytest
from pathlib import Path
from hwpconv import HwpxParser, MarkdownConverter, Document


# 테스트 파일 경로
FIXTURES_DIR = Path(__file__).parent / 'fixtures'


class TestHwpxParser:
    """HWPX 파서 테스트"""
    
    def test_can_parse(self):
        """파싱 가능 여부 확인"""
        assert HwpxParser.can_parse('test.hwpx') is True
        assert HwpxParser.can_parse('test.HWPX') is True
        assert HwpxParser.can_parse('test.hwp') is False
        assert HwpxParser.can_parse('test.docx') is False
    
    def test_quick_extract_nonexistent(self):
        """존재하지 않는 파일 빠른 추출"""
        result = HwpxParser.quick_extract('nonexistent.hwpx')
        assert result == ''
    
    def test_parse_creates_document(self):
        """파싱 결과가 Document 객체인지 확인"""
        # 실제 HWPX 파일이 있으면 테스트
        hwpx_file = FIXTURES_DIR / 'sample.hwpx'
        if hwpx_file.exists():
            parser = HwpxParser()
            doc = parser.parse(str(hwpx_file))
            assert isinstance(doc, Document)


class TestHwpxParserWithFixture:
    """실제 HWPX 파일을 사용한 테스트"""
    
    @pytest.fixture
    def sample_hwpx(self):
        """샘플 HWPX 파일 경로"""
        # 프로젝트 루트의 테스트 파일 사용
        test_file = Path(__file__).parent.parent.parent / '회생_기각결정문_반박내용.hwpx'
        if test_file.exists():
            return test_file
        pytest.skip("테스트 파일 없음")
    
    def test_parse_hwpx(self, sample_hwpx):
        """HWPX 파일 파싱"""
        parser = HwpxParser()
        doc = parser.parse(str(sample_hwpx))
        
        assert isinstance(doc, Document)
        assert len(doc.sections) > 0
        assert doc.total_paragraph_count > 0
    
    def test_quick_extract(self, sample_hwpx):
        """빠른 텍스트 추출"""
        text = HwpxParser.quick_extract(str(sample_hwpx))
        assert isinstance(text, str)
    
    def test_convert_to_markdown(self, sample_hwpx):
        """Markdown 변환"""
        parser = HwpxParser()
        doc = parser.parse(str(sample_hwpx))
        
        converter = MarkdownConverter()
        md = converter.convert(doc)
        
        assert isinstance(md, str)
        assert len(md) > 0
