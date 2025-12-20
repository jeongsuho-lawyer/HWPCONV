"""
CLI 엔트리포인트
"""

import argparse
import sys
from pathlib import Path
from typing import Optional


def main():
    parser = argparse.ArgumentParser(
        prog='hwpconv',
        description='HWP/HWPX → Markdown/HTML 변환기'
    )
    parser.add_argument('input', help='입력 파일 (.hwp, .hwpx)')
    parser.add_argument('-o', '--output', help='출력 파일')
    parser.add_argument('-f', '--format', choices=['md', 'html', 'txt'], 
                        default='md', help='출력 포맷 (기본: md)')
    parser.add_argument('--quick', action='store_true', 
                        help='빠른 텍스트 추출 (Preview 활용)')
    parser.add_argument('--no-images', action='store_true',
                        help='이미지 포함하지 않음')
    parser.add_argument('--analyze-images', action='store_true',
                        help='이미지 내용 분석 (Gemini Vision API 사용)')
    parser.add_argument('--api-key', help='Gemini API 키 (또는 GOOGLE_API_KEY 환경변수)')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f'Error: {input_path} not found', file=sys.stderr)
        sys.exit(1)
    
    # 파서/변환기 임포트 (지연 로딩으로 시작 시간 단축)
    from .parsers.hwpx import HwpxParser
    from .parsers.hwp import HwpParser
    from .converters.markdown import MarkdownConverter
    from .converters.html import HtmlConverter
    
    # 파서 선택
    ext = input_path.suffix.lower()
    
    if ext == '.hwpx':
        if args.quick:
            result = HwpxParser.quick_extract(str(input_path))
            _output(result, args.output)
            return
        doc = HwpxParser().parse(str(input_path))
    elif ext == '.hwp':
        if args.quick:
            result = HwpParser.quick_extract(str(input_path))
            _output(result, args.output)
            return
        doc = HwpParser().parse(str(input_path))
    else:
        print(f'Error: Unsupported format {ext}', file=sys.stderr)
        sys.exit(1)
    
    # 이미지 분석 (옵션) - 파서에서 이미 분석함
    if args.analyze_images and doc.images:
        print(f'이미지 {len(doc.images)}개 분석 완료 (파싱 단계에서 처리됨)', file=sys.stderr)
    
    # 변환기 선택
    include_images = not args.no_images
    
    if args.format == 'html':
        result = HtmlConverter(include_images=include_images).convert(doc)
    elif args.format == 'txt':
        result = doc.text
    else:  # md
        result = MarkdownConverter(include_images=include_images).convert(doc)
    
    _output(result, args.output)


def _output(content: str, output_path: Optional[str]) -> None:
    """결과 출력"""
    if output_path:
        Path(output_path).write_text(content, encoding='utf-8')
        print(f'Saved to {output_path}', file=sys.stderr)
    else:
        print(content)


if __name__ == '__main__':
    main()
