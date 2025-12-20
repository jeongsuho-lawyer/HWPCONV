# hwpconv

한글 문서(.hwp, .hwpx)를 Markdown/HTML/Text로 변환하는 Python 라이브러리 및 CLI

## 주요 기능

- ✅ **HWPX 파서**: ZIP+XML 형식 완벽 지원
- ✅ **HWP 파서**: OLE+zlib 바이너리 형식 지원
- ✅ **Markdown 변환**: 서식, 표, 각주 포함
- ✅ **HTML 변환**: 스타일 포함 HTML 출력
- ✅ **이미지 추출**: Base64 인라인 포함 (Claude 등 AI가 직접 볼 수 있음)
- ✅ **이미지 분석 (선택)**: Gemini Vision API로 이미지 내용 자동 설명
- ✅ **웹 UI**: 로컬 서버 기반 드래그앤드롭 변환기
- ✅ **빠른 추출**: Preview 텍스트 활용

## 설치

```bash
# 기본 설치
pip install -e .

# 개발 의존성 포함
pip install -e ".[dev]"

# 이미지 분석 기능 (선택)
pip install google-generativeai
```

## CLI 사용법

```bash
# 기본 변환 (Markdown)
hwpconv document.hwpx -o output.md
hwpconv document.hwp -o output.md

# HTML 변환
hwpconv document.hwpx -f html -o output.html

# 텍스트만 추출
hwpconv document.hwp -f txt

# 빠른 텍스트 추출 (Preview 활용)
hwpconv document.hwpx --quick

# 이미지 제외
hwpconv document.hwpx --no-images -o output.md

# 이미지 분석 포함 (Gemini API 필요)
export GOOGLE_API_KEY="your-api-key"
hwpconv document.hwpx --analyze-images -o output.md
```

## Python API

```python
from hwpconv import HwpxParser, HwpParser, MarkdownConverter, HtmlConverter

# HWPX 파싱
doc = HwpxParser().parse('document.hwpx')

# HWP 파싱
doc = HwpParser().parse('document.hwp')

# Markdown 변환
md = MarkdownConverter().convert(doc)

# HTML 변환
html = HtmlConverter().convert(doc)

# 문서 정보
print(f"문단 수: {doc.total_paragraph_count}")
print(f"표 개수: {doc.total_table_count}")  
print(f"이미지 수: {doc.total_image_count}")

# 빠른 텍스트 추출
text = HwpxParser.quick_extract('document.hwpx')
text = HwpParser.quick_extract('document.hwp')
```

## 데스크톱 GUI (Windows 네이티브)

```bash
# GUI 앱 실행 (브라우저 필요 없음)
python -m hwpconv.gui

# 또는 설치 후
hwpconv-gui
```

GUI 기능:
- 파일 추가 버튼으로 HWP/HWPX 선택
- Markdown/HTML/Text 형식 선택
- 이미지 포함 옵션
- 일괄 변환 후 폴더 열기

## 웹 UI (선택적, Flask 필요)

```bash
# Flask 서버 실행
pip install flask
python -m hwpconv.server

# 특정 포트로 실행
python -m hwpconv.server -p 8080
```

## 이미지 분석 (선택적 기능)

이미지가 포함된 문서를 변환할 때, AI가 이미지 내용을 이해할 수 있도록 설명을 자동 생성합니다.

```python
from hwpconv import HwpxParser, MarkdownConverter
from hwpconv.image_analyzer import analyze_document_images

# 문서 파싱
doc = HwpxParser().parse('document.hwpx')

# 이미지 분석 (Gemini Vision API)
analyze_document_images(doc, provider='gemini', api_key='your-key')

# 변환 (이미지 설명 포함)
md = MarkdownConverter(include_images=True).convert(doc)
```

**지원 AI 제공자:**
- `gemini` (기본): Google Gemini Vision (무료 티어 제공)
- `openai`: GPT-4 Vision
- `anthropic`: Claude Vision

## 지원 형식

| 입력 | 출력 | 상태 |
|------|------|------|
| .hwpx (ZIP+XML) | Markdown | ✅ |
| .hwpx | HTML | ✅ |
| .hwpx | Text | ✅ |
| .hwp (OLE) | Markdown | ✅ |
| .hwp | HTML | ✅ |
| .hwp | Text | ✅ |

## 지원 요소

- ✅ 일반 텍스트
- ✅ 서식 (굵게, 기울임, 밑줄, 취소선)
- ✅ 제목 레벨 (H1-H6)
- ✅ 표 (colspan, rowspan)
- ✅ 각주/미주
- ✅ 이미지 (Base64 인라인)
- ⚠️ 글머리 기호/번호 (부분 지원)

## 라이선스

MIT License

### 저작권 고지

본 소프트웨어는 한글과컴퓨터 한/글 문서 형식(.hwp, .hwpx)의 구조를 분석하여 
텍스트 변환 기능을 제공합니다. "한/글" 및 "HWP"는 (주)한글과컴퓨터의 등록상표입니다.