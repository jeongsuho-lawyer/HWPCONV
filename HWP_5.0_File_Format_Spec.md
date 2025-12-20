# 한글 문서 파일 구조 5.0 (HWP Document File Formats 5.0)

**revision 1.3:20181108**

---

## 저작권

(주)한글과컴퓨터는 문서 형식의 개방성과 표준화에 적극 찬성합니다. 본 문서를 누구나 열람, 복사, 배포, 게재 및 사용할 수 있습니다. 다만, 배포는 원 내용이 일체 수정되지 않은 원본 또는 복사본으로 제한됩니다.

개발 결과물에 **"본 제품은 한글과컴퓨터의 한글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다."**라고 기재해야 합니다.

---

## 1. 개요

- 기본 확장자: `.HWP`
- 압축: **zlib** 사용
- 파일 구조: **Windows Compound File**
- 문자 코드: **ISO-10646 (UTF-16LE)**
- 기본 단위: **1/7200인치** (HWPUNIT)

---

## 2. 자료형

| 자료형 | 크기(byte) | 설명 |
|--------|------------|------|
| BYTE | 1 | 0~255 |
| WORD | 2 | unsigned 16-bit |
| DWORD | 4 | unsigned 32-bit |
| WCHAR | 2 | UTF-16LE 문자 |
| HWPUNIT | 4 | 1/7200인치 단위 |
| SHWPUNIT | 4 | signed HWPUNIT |
| COLORREF | 4 | 0x00bbggrr |
| UINT8/16/32 | 1/2/4 | unsigned |
| INT8/16/32 | 1/2/4 | signed |

---

## 3. 파일 구조

### 3.1 전체 구조

| 설명 | 스트림명 | 압축 | 레코드구조 |
|------|----------|------|-----------|
| 파일 인식 정보 | FileHeader (256B) | - | - |
| 문서 정보 | DocInfo | ✓ | ✓ |
| 본문 | BodyText/Section0,1,... | ✓ | ✓ |
| 문서 요약 | \005HwpSummaryInformation | - | - |
| 바이너리 데이터 | BinData/BinaryData0,1,... | ✓ | - |
| 미리보기 텍스트 | PrvText | - | - |
| 미리보기 이미지 | PrvImage | - | - |
| 문서 옵션 | DocOptions/* | - | - |
| 스크립트 | Scripts/* | - | - |
| XML 템플릿 | XMLTemplate/* | - | - |
| 문서 이력 | DocHistory/VersionLog0,1,... | ✓ | ✓ |

### 3.2 FileHeader (256 bytes)

```
Offset  Size  Description
0x00    32    Signature: "HWP Document File"
0x20    4     Version (0xMMnnPPrr)
0x24    4     속성 플래그
0x28    4     CCL/공공누리 플래그
0x2C    4     EncryptVersion (0~4)
0x30    1     KOGL 국가코드 (6=KOR, 15=US)
0x31    207   Reserved
```

**속성 플래그 (0x24):**
| Bit | 설명 |
|-----|------|
| 0 | 압축 여부 |
| 1 | 암호 설정 |
| 2 | 배포용 문서 |
| 3 | 스크립트 저장 |
| 4 | DRM 보안 |
| 5 | XMLTemplate 존재 |
| 6 | 문서 이력 존재 |
| 7 | 전자 서명 존재 |

---

## 4. 데이터 레코드 구조

### 4.1 레코드 헤더 (32 bits)

```
┌──────────┬──────────┬──────────────┐
│ Tag ID   │ Level    │ Size         │
│ (10 bit) │ (10 bit) │ (12 bit)     │
└──────────┴──────────┴──────────────┘
```

- **Tag ID**: 0x010~0x1FF (한글 내부용), 0x200~0x3FF (외부용)
- **Level**: 레코드 계층 depth
- **Size**: 데이터 길이 (0xFFF 이상이면 추가 DWORD)

### 4.2 문서 정보 레코드 (DocInfo)

| Tag ID | Value | 설명 |
|--------|-------|------|
| HWPTAG_DOCUMENT_PROPERTIES | 0x010 | 문서 속성 |
| HWPTAG_ID_MAPPINGS | 0x011 | 아이디 매핑 |
| HWPTAG_BIN_DATA | 0x012 | 바이너리 데이터 |
| HWPTAG_FACE_NAME | 0x013 | 글꼴 |
| HWPTAG_BORDER_FILL | 0x014 | 테두리/배경 |
| HWPTAG_CHAR_SHAPE | 0x015 | 글자 모양 |
| HWPTAG_TAB_DEF | 0x016 | 탭 정의 |
| HWPTAG_NUMBERING | 0x017 | 번호 정의 |
| HWPTAG_BULLET | 0x018 | 불릿 정의 |
| HWPTAG_PARA_SHAPE | 0x019 | 문단 모양 |
| HWPTAG_STYLE | 0x01A | 스타일 |
| HWPTAG_DOC_DATA | 0x01B | 문서 임의 데이터 |
| HWPTAG_DISTRIBUTE_DOC_DATA | 0x01C | 배포용 문서 |
| HWPTAG_COMPATIBLE_DOCUMENT | 0x01E | 호환 문서 |
| HWPTAG_LAYOUT_COMPATIBILITY | 0x01F | 레이아웃 호환성 |

### 4.3 본문 레코드 (BodyText)

| Tag ID | Value | 설명 |
|--------|-------|------|
| HWPTAG_PARA_HEADER | 0x042 | 문단 헤더 |
| HWPTAG_PARA_TEXT | 0x043 | 문단 텍스트 |
| HWPTAG_PARA_CHAR_SHAPE | 0x044 | 문단 글자모양 |
| HWPTAG_PARA_LINE_SEG | 0x045 | 문단 레이아웃 |
| HWPTAG_PARA_RANGE_TAG | 0x046 | 문단 영역태그 |
| HWPTAG_CTRL_HEADER | 0x047 | 컨트롤 헤더 |
| HWPTAG_LIST_HEADER | 0x048 | 문단리스트 헤더 |
| HWPTAG_PAGE_DEF | 0x049 | 용지 설정 |
| HWPTAG_FOOTNOTE_SHAPE | 0x04A | 각주/미주 모양 |
| HWPTAG_PAGE_BORDER_FILL | 0x04B | 쪽 테두리/배경 |
| HWPTAG_SHAPE_COMPONENT | 0x04C | 개체 |
| HWPTAG_TABLE | 0x04D | 표 |
| HWPTAG_SHAPE_COMPONENT_LINE | 0x04E | 직선 |
| HWPTAG_SHAPE_COMPONENT_RECTANGLE | 0x04F | 사각형 |
| HWPTAG_SHAPE_COMPONENT_ELLIPSE | 0x050 | 타원 |
| HWPTAG_SHAPE_COMPONENT_ARC | 0x051 | 호 |
| HWPTAG_SHAPE_COMPONENT_POLYGON | 0x052 | 다각형 |
| HWPTAG_SHAPE_COMPONENT_CURVE | 0x053 | 곡선 |
| HWPTAG_SHAPE_COMPONENT_OLE | 0x054 | OLE |
| HWPTAG_SHAPE_COMPONENT_PICTURE | 0x055 | 그림 |
| HWPTAG_SHAPE_COMPONENT_CONTAINER | 0x056 | 묶음개체 |
| HWPTAG_CTRL_DATA | 0x057 | 컨트롤 데이터 |
| HWPTAG_EQEDIT | 0x058 | 수식 |

---

## 5. 제어 문자 (컨트롤)

| 코드 | 설명 | 형식 | 크기 |
|------|------|------|------|
| 0 | unusable | char | 1 |
| 2 | 구역/단 정의 | extended | 8 |
| 3 | 필드 시작 | extended | 8 |
| 4 | 필드 끝 | inline | 8 |
| 9 | 탭 | inline | 8 |
| 10 | 줄바꿈 (line break) | char | 1 |
| 11 | 그리기개체/표 | extended | 8 |
| 13 | 문단끝 (para break) | char | 1 |
| 15 | 숨은설명 | extended | 8 |
| 16 | 머리말/꼬리말 | extended | 8 |
| 17 | 각주/미주 | extended | 8 |
| 18 | 자동번호 | extended | 8 |
| 21 | 페이지 컨트롤 | extended | 8 |
| 22 | 책갈피/찾아보기 | extended | 8 |
| 23 | 덧말/글자겹침 | extended | 8 |
| 24 | 하이픈 | char | 1 |
| 30 | 묶음 빈칸 | char | 1 |
| 31 | 고정폭 빈칸 | char | 1 |

---

## 6. 주요 구조체

### 6.1 문서 속성 (HWPTAG_DOCUMENT_PROPERTIES)

| Offset | Size | Description |
|--------|------|-------------|
| 0 | 2 | 구역 개수 |
| 2 | 2 | 페이지 시작 번호 |
| 4 | 2 | 각주 시작 번호 |
| 6 | 2 | 미주 시작 번호 |
| 8 | 2 | 그림 시작 번호 |
| 10 | 2 | 표 시작 번호 |
| 12 | 2 | 수식 시작 번호 |
| 14 | 4 | 리스트 ID |
| 18 | 4 | 문단 ID |
| 22 | 4 | 문단 내 위치 |
| **Total** | **26** | |

### 6.2 글자 모양 (HWPTAG_CHAR_SHAPE)

| Offset | Size | Description |
|--------|------|-------------|
| 0 | 14 | 언어별 글꼴 ID (7개) |
| 14 | 7 | 언어별 장평 |
| 21 | 7 | 언어별 자간 |
| 28 | 7 | 언어별 상대크기 |
| 35 | 7 | 언어별 글자위치 |
| 42 | 4 | 기준 크기 |
| 46 | 4 | 속성 |
| 50 | 2 | 그림자 간격 |
| 52 | 4 | 글자 색 |
| 56 | 4 | 밑줄 색 |
| 60 | 4 | 음영 색 |
| 64 | 4 | 그림자 색 |
| 68 | 2 | 테두리/배경 ID |
| 70 | 4 | 취소선 색 |
| **Total** | **72** | |

**속성 비트:**
- bit 0: 기울임
- bit 1: 진하게
- bit 2-3: 밑줄 종류
- bit 4-7: 밑줄 모양
- bit 8-10: 외곽선 종류
- bit 11-12: 그림자 종류
- bit 13: 양각
- bit 14: 음각
- bit 15: 위첨자
- bit 16: 아래첨자
- bit 18-20: 취소선
- bit 21-24: 강조점
- bit 30: Kerning

### 6.3 문단 모양 (HWPTAG_PARA_SHAPE)

| Offset | Size | Description |
|--------|------|-------------|
| 0 | 4 | 속성1 |
| 4 | 4 | 왼쪽 여백 |
| 8 | 4 | 오른쪽 여백 |
| 12 | 4 | 들여쓰기/내어쓰기 |
| 16 | 4 | 문단 간격 위 |
| 20 | 4 | 문단 간격 아래 |
| 24 | 4 | 줄 간격 (구버전) |
| 28 | 2 | 탭 정의 ID |
| 30 | 2 | 번호/글머리표 ID |
| 32 | 2 | 테두리/배경 ID |
| 34 | 8 | 문단 테두리 간격 |
| 42 | 4 | 속성2 |
| 46 | 4 | 속성3 |
| 50 | 4 | 줄 간격 (신버전) |
| **Total** | **54** | |

### 6.4 문단 헤더 (HWPTAG_PARA_HEADER)

| Offset | Size | Description |
|--------|------|-------------|
| 0 | 4 | 텍스트 문자 수 |
| 4 | 4 | 컨트롤 마스크 |
| 8 | 2 | 문단 모양 ID |
| 10 | 1 | 스타일 ID |
| 11 | 1 | 단 나누기 종류 |
| 12 | 2 | 글자모양 정보 수 |
| 14 | 2 | 영역태그 정보 수 |
| 16 | 2 | 줄 정보 수 |
| 18 | 4 | 문단 Instance ID |
| 22 | 2 | 변경추적 병합 여부 |
| **Total** | **24** | |

---

## 7. 컨트롤 ID

### 7.1 개체 컨트롤

```c
#define MAKE_4CHID(a,b,c,d) (((a)<<24)|((b)<<16)|((c)<<8)|(d))

표:      MAKE_4CHID('t','b','l',' ')  // 0x74626C20
선:      MAKE_4CHID('$','l','i','n')  // 0x246C696E
사각형:  MAKE_4CHID('$','r','e','c')  // 0x24726563
타원:    MAKE_4CHID('$','e','l','l')  // 0x24656C6C
호:      MAKE_4CHID('$','a','r','c')  // 0x24617263
다각형:  MAKE_4CHID('$','p','o','l')  // 0x24706F6C
곡선:    MAKE_4CHID('$','c','u','r')  // 0x24637572
수식:    MAKE_4CHID('e','q','e','d')  // 0x65716564
그림:    MAKE_4CHID('$','p','i','c')  // 0x24706963
OLE:     MAKE_4CHID('$','o','l','e')  // 0x246F6C65
묶음:    MAKE_4CHID('$','c','o','n')  // 0x24636F6E
```

### 7.2 비개체 컨트롤

```c
구역정의:    MAKE_4CHID('s','e','c','d')
단정의:      MAKE_4CHID('c','o','l','d')
머리말:      MAKE_4CHID('h','e','a','d')
꼬리말:      MAKE_4CHID('f','o','o','t')
각주:        MAKE_4CHID('f','n',' ',' ')
미주:        MAKE_4CHID('e','n',' ',' ')
자동번호:    MAKE_4CHID('a','t','n','o')
새번호:      MAKE_4CHID('n','w','n','o')
감추기:      MAKE_4CHID('p','g','h','d')
홀짝조정:    MAKE_4CHID('p','g','c','t')
쪽번호위치:  MAKE_4CHID('p','g','n','p')
찾아보기:    MAKE_4CHID('i','d','x','m')
책갈피:      MAKE_4CHID('b','o','k','m')
글자겹침:    MAKE_4CHID('t','c','p','s')
덧말:        MAKE_4CHID('t','d','u','t')
숨은설명:    MAKE_4CHID('t','c','m','t')
```

### 7.3 필드 컨트롤

```c
FIELD_UNKNOWN:      MAKE_4CHID('%','u','n','k')
FIELD_DATE:         MAKE_4CHID('%','d','t','e')
FIELD_DOCDATE:      MAKE_4CHID('%','d','d','t')
FIELD_PATH:         MAKE_4CHID('%','p','a','t')
FIELD_BOOKMARK:     MAKE_4CHID('%','b','m','k')
FIELD_MAILMERGE:    MAKE_4CHID('%','m','m','g')
FIELD_CROSSREF:     MAKE_4CHID('%','x','r','f')
FIELD_FORMULA:      MAKE_4CHID('%','f','m','u')
FIELD_CLICKHERE:    MAKE_4CHID('%','c','l','k')
FIELD_SUMMARY:      MAKE_4CHID('%','s','m','r')
FIELD_USERINFO:     MAKE_4CHID('%','u','s','r')
FIELD_HYPERLINK:    MAKE_4CHID('%','h','l','k')
FIELD_MEMO:         MAKE_4CHID('%','%','m','e')
```

---

## 8. 테두리선 종류/굵기

### 테두리선 종류

| 값 | 설명 |
|----|------|
| 0 | 실선 |
| 1 | 긴 점선 |
| 2 | 점선 |
| 3 | -.-.-. |
| 4 | -..-.. |
| 5 | 긴 대시 |
| 6 | 큰 원 |
| 7 | 2중선 |
| 8 | 가는선+굵은선 |
| 9 | 굵은선+가는선 |
| 10 | 3중선 |
| 11 | 물결 |
| 12 | 물결 2중선 |
| 13-16 | 3D 효과 |

### 테두리선 굵기

| 값 | mm | 값 | mm |
|----|-----|----|----|
| 0 | 0.1 | 8 | 0.6 |
| 1 | 0.12 | 9 | 0.7 |
| 2 | 0.15 | 10 | 1.0 |
| 3 | 0.2 | 11 | 1.5 |
| 4 | 0.25 | 12 | 2.0 |
| 5 | 0.3 | 13 | 3.0 |
| 6 | 0.4 | 14 | 4.0 |
| 7 | 0.5 | 15 | 5.0 |

---

## 9. 번호 형식

| 값 | 형식 |
|----|------|
| 0 | 1, 2, 3 |
| 1 | ①, ②, ③ |
| 2 | I, II, III |
| 3 | i, ii, iii |
| 4 | A, B, C |
| 5 | a, b, c |
| 6 | Ⓐ, Ⓑ, Ⓒ |
| 7 | ⓐ, ⓑ, ⓒ |
| 8 | 가, 나, 다 |
| 9 | ㉮, ㉯, ㉰ |
| 10 | ㄱ, ㄴ, ㄷ |
| 11 | ㉠, ㉡, ㉢ |
| 12 | 일, 이, 삼 |
| 13 | 一, 二, 三 |
| 14 | ㊀, ㊁, ㊂ |

---

## 10. 버전 이력

| 버전 | 날짜 | 변경내용 |
|------|------|---------|
| 1.0 | 2010-07-01 | 최초 공개 |
| 1.1 | 2011-01-24 | 저작권 수정, 오타 수정 |
| 1.2 | 2014-10-09 | 파트별 구성, 상세 내용 추가 |
| 1.3 | 2018-11-08 | 참고문헌, 차트, 동영상, 그림효과 추가 |

---

**발행:** (주)한글과컴퓨터  
**주소:** 경기도 성남시 분당구 대왕판교로 644번길 49 한컴타워
