# HWP2MD

[![GitHub release](https://img.shields.io/github/v/release/jeongsuho-lawyer/HWPCONV)](https://github.com/jeongsuho-lawyer/HWPCONV/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

한글 문서(HWP/HWPX)를 Markdown 또는 HTML로 변환하는 Windows 데스크톱 프로그램입니다.

## 주요 기능

- **HWP/HWPX 변환**: 한글 문서를 Markdown(.md) 또는 HTML(.html)로 변환
- **드래그 앤 드롭**: 파일이나 폴더를 창에 끌어다 놓으면 자동 변환
- **이미지 분석**: Google Gemini API를 활용하여 문서 내 이미지 내용을 텍스트로 설명
- **표 변환**: 문서 내 표를 마크다운/HTML 테이블로 변환
- **일괄 처리**: 여러 파일을 한 번에 변환하고 저장
- **파일별 형식 지정**: 각 파일마다 MD/HTML 형식을 개별 지정 가능

## 사용 방법

### 기본 사용

1. `HWP2MD.exe` 실행
2. 상단에서 변환 형식 선택 (`마크다운` 또는 `HTML`)
3. HWP/HWPX 파일을 창으로 드래그 앤 드롭
4. 변환 완료 후 `저장` 또는 `전체 저장` 클릭

### 이미지 분석 기능

문서 내 이미지를 AI가 분석하여 설명 텍스트를 생성합니다.

1. 우측 상단 `설정` 클릭
2. Google Gemini API 키 입력 후 저장
3. `이미지 분석` 옵션을 `사용`으로 변경
4. 파일 추가 시 이미지가 자동 분석됨

API 키는 [Google AI Studio](https://aistudio.google.com/app/apikey)에서 **무료로 발급**받을 수 있습니다.

### 모델 선택

설정에서 이미지 분석에 사용할 Gemini 모델을 선택할 수 있습니다:

| 모델 | 특징 | 무료 제한 |
|------|------|-----------|
| Gemini 3 Flash | 최신, 고성능 | 분당 10개 |
| Gemini 2.5 Flash | 균형잡힌 성능 | 분당 15개 |
| Gemini 2.5 Flash Lite | 초고속 | 분당 30개 |
| Gemini 2.0 Flash | 안정적 | 분당 15개 |

새로운 모델이 출시되면 "직접 입력"으로 모델 ID를 입력하여 바로 사용할 수 있습니다.

### 출력 예시

**Markdown 출력:**
```markdown
# 문서 제목

본문 내용입니다.

| 항목 | 설명 |
|------|------|
| A | 내용1 |
| B | 내용2 |

> 🖼️ **[이미지]**: 회사 로고 이미지
```

**이미지 분석 미사용 시:**
```markdown
> 🖼️ **[이미지]**: *(이미지 생략)*
```

## 설치 및 실행

### 실행 파일 (권장)

`HWP2MD.exe` 파일을 다운로드하여 바로 실행하면 됩니다. 별도 설치가 필요 없습니다.

### 소스에서 실행

```bash
pip install -r requirements.txt
python run_gui.py
```

### EXE 직접 빌드

```bash
pip install -r requirements.txt
pyinstaller HwpConverterPro.spec
```

빌드 결과: `dist/HWP2MD.exe`

> **참고**: `requirements.txt`에 의존성 버전이 고정되어 있어 동일한 빌드 환경을 재현할 수 있습니다.

## 시스템 요구사항

- Windows 10 / 11
- 인터넷 연결 (이미지 분석 기능 사용 시)

## 저장 위치

- 변환된 파일: EXE 실행 폴더 내 `HWPCONV_Output` 폴더
- 설정 파일: `%APPDATA%\HwpConverter\config.json`

## 지원 요소

| 요소 | 지원 |
|------|------|
| 텍스트 | ✅ |
| 서식 (굵게, 기울임, 밑줄) | ✅ |
| 제목 (H1~H6) | ✅ |
| 표 | ✅ |
| 이미지 | ✅ |
| 각주/미주 | ✅ |

## 개인정보 보호

- API 키는 사용자 PC에만 로컬 저장됨
- 이미지 분석 사용 시에만 해당 이미지가 Google 서버로 전송됨
- 문서 원본은 외부로 전송되지 않음

## 제작 정보

- **제작**: 법무법인 르네상스 정수호 변호사
- **연락처**: shj@lawren.co.kr
- **버전**: 1.0.1

## 변경 이력

### v1.0.1
- 특수 문자(UTF-16 서로게이트) 포함 시 빈 파일 생성되는 버그 수정
- HWP/HWPX 파서에서 비정상 유니코드 문자 자동 필터링 추가

### v1.0.0
- 최초 릴리스

## 라이선스

MIT License

본 소프트웨어는 한글과컴퓨터의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발되었습니다.
