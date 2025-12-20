"""
HWP 파서 테스트
"""

import pytest
from pathlib import Path
from hwpconv import HwpParser, MarkdownConverter, Document


FIXTURES_DIR = Path(__file__).parent / 'fixtures'


class TestHwpParser:
    """HWP 파서 테스트"""
    
    def test_can_parse(self):
        """파싱 가능 여부 확인"""
        assert HwpParser.can_parse('test.hwp') is True
        assert HwpParser.can_parse('test.HWP') is True
        assert HwpParser.can_parse('test.hwpx') is False
        assert HwpParser.can_parse('test.doc') is False
    
    def test_quick_extract_nonexistent(self):
        """존재하지 않는 파일 빠른 추출"""
        result = HwpParser.quick_extract('nonexistent.hwp')
        assert result == ''


class TestHwpParserWithFixture:
    """실제 HWP 파일을 사용한 테스트"""
    
    @pytest.fixture
    def sample_hwp(self):
        """샘플 HWP 파일 경로"""
        test_file = Path(__file__).parent.parent.parent / '비영리법인 설립신청 및 업무서식 안내-공개용.hwp'
        if test_file.exists():
            return test_file
        pytest.skip("테스트 파일 없음")
    
    def test_parse_hwp(self, sample_hwp):
        """HWP 파일 파싱"""
        parser = HwpParser()
        doc = parser.parse(str(sample_hwp))
        
        assert isinstance(doc, Document)
        assert len(doc.sections) > 0
    
    def test_quick_extract(self, sample_hwp):
        """빠른 텍스트 추출"""
        text = HwpParser.quick_extract(str(sample_hwp))
        assert isinstance(text, str)
        assert len(text) > 0
    
    def test_convert_to_markdown(self, sample_hwp):
        """Markdown 변환"""
        parser = HwpParser()
        doc = parser.parse(str(sample_hwp))
        
        converter = MarkdownConverter()
        md = converter.convert(doc)
        
        assert isinstance(md, str)
