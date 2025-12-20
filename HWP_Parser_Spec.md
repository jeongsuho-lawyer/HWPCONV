# HWP 5.0 Parser Specification for Markdown Conversion

## 1. 파일 구조 개요

```
HWP 파일 (.hwp)
├── FileHeader (256 bytes, 비압축)
├── DocInfo (압축, 레코드 구조)
├── BodyText/
│   ├── Section0 (압축, 레코드 구조)
│   ├── Section1
│   └── ...
├── BinData/ (이미지 등 바이너리)
│   ├── BIN0001.jpg
│   └── ...
└── PrvText (미리보기 텍스트, 비압축)
```

- **파일 형식**: OLE2 Compound File (python: `olefile`)
- **압축**: zlib (`decompress()`)
- **문자 인코딩**: UTF-16LE
- **좌표 단위**: HWPUNIT = 1/7200 inch

---

## 2. FileHeader 파싱 (256 bytes)

```python
def parse_file_header(data: bytes) -> dict:
    signature = data[0:32]  # b"HWP Document File" + padding
    if not signature.startswith(b"HWP Document File"):
        raise ValueError("Not a HWP file")
    
    version = struct.unpack('<I', data[32:36])[0]
    # version format: 0xMMnnPPrr (Major.minor.Patch.revision)
    # 예: 0x05010106 = 5.1.1.6
    
    flags = struct.unpack('<I', data[36:40])[0]
    
    return {
        'version': version,
        'compressed': bool(flags & 0x01),      # bit 0
        'encrypted': bool(flags & 0x02),       # bit 1
        'distributed': bool(flags & 0x04),     # bit 2: 배포용 문서
        'has_script': bool(flags & 0x08),      # bit 3
        'has_drm': bool(flags & 0x10),         # bit 4
        'has_xml_template': bool(flags & 0x20),# bit 5
        'has_history': bool(flags & 0x40),     # bit 6
        'has_signature': bool(flags & 0x80),   # bit 7
    }
```

---

## 3. 레코드 구조 파싱

모든 DocInfo, BodyText 스트림은 연속된 레코드로 구성.

### 레코드 헤더 (4 bytes)

```
┌────────────┬────────────┬──────────────┐
│  Tag ID    │   Level    │    Size      │
│  (10 bit)  │  (10 bit)  │  (12 bit)    │
└────────────┴────────────┴──────────────┘
```

```python
def parse_records(data: bytes) -> list:
    records = []
    pos = 0
    while pos < len(data):
        if pos + 4 > len(data):
            break
        header = struct.unpack('<I', data[pos:pos+4])[0]
        tag_id = header & 0x3FF           # 하위 10비트
        level = (header >> 10) & 0x3FF    # 중간 10비트
        size = (header >> 20) & 0xFFF     # 상위 12비트
        pos += 4
        
        # size가 0xFFF면 다음 4바이트가 실제 크기
        if size == 0xFFF:
            size = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
        
        payload = data[pos:pos+size]
        pos += size
        
        records.append({
            'tag_id': tag_id,
            'level': level,
            'size': size,
            'data': payload
        })
    return records
```

---

## 4. Tag ID 정의

```python
HWPTAG_BEGIN = 0x010

# DocInfo 레코드
HWPTAG_DOCUMENT_PROPERTIES    = HWPTAG_BEGIN + 0   # 0x010
HWPTAG_ID_MAPPINGS            = HWPTAG_BEGIN + 1   # 0x011
HWPTAG_BIN_DATA               = HWPTAG_BEGIN + 2   # 0x012
HWPTAG_FACE_NAME              = HWPTAG_BEGIN + 3   # 0x013
HWPTAG_BORDER_FILL            = HWPTAG_BEGIN + 4   # 0x014
HWPTAG_CHAR_SHAPE             = HWPTAG_BEGIN + 5   # 0x015
HWPTAG_TAB_DEF                = HWPTAG_BEGIN + 6   # 0x016
HWPTAG_NUMBERING              = HWPTAG_BEGIN + 7   # 0x017
HWPTAG_BULLET                 = HWPTAG_BEGIN + 8   # 0x018
HWPTAG_PARA_SHAPE             = HWPTAG_BEGIN + 9   # 0x019
HWPTAG_STYLE                  = HWPTAG_BEGIN + 10  # 0x01A
HWPTAG_DOC_DATA               = HWPTAG_BEGIN + 11  # 0x01B
HWPTAG_DISTRIBUTE_DOC_DATA    = HWPTAG_BEGIN + 12  # 0x01C
HWPTAG_COMPATIBLE_DOCUMENT    = HWPTAG_BEGIN + 14  # 0x01E
HWPTAG_LAYOUT_COMPATIBILITY   = HWPTAG_BEGIN + 15  # 0x01F

# BodyText 레코드
HWPTAG_PARA_HEADER            = HWPTAG_BEGIN + 50  # 0x042
HWPTAG_PARA_TEXT              = HWPTAG_BEGIN + 51  # 0x043
HWPTAG_PARA_CHAR_SHAPE        = HWPTAG_BEGIN + 52  # 0x044
HWPTAG_PARA_LINE_SEG          = HWPTAG_BEGIN + 53  # 0x045
HWPTAG_PARA_RANGE_TAG         = HWPTAG_BEGIN + 54  # 0x046
HWPTAG_CTRL_HEADER            = HWPTAG_BEGIN + 55  # 0x047
HWPTAG_LIST_HEADER            = HWPTAG_BEGIN + 56  # 0x048
HWPTAG_PAGE_DEF               = HWPTAG_BEGIN + 57  # 0x049
HWPTAG_FOOTNOTE_SHAPE         = HWPTAG_BEGIN + 58  # 0x04A
HWPTAG_PAGE_BORDER_FILL       = HWPTAG_BEGIN + 59  # 0x04B
HWPTAG_SHAPE_COMPONENT        = HWPTAG_BEGIN + 60  # 0x04C
HWPTAG_TABLE                  = HWPTAG_BEGIN + 61  # 0x04D
HWPTAG_SHAPE_COMPONENT_LINE   = HWPTAG_BEGIN + 62  # 0x04E
HWPTAG_SHAPE_COMPONENT_RECT   = HWPTAG_BEGIN + 63  # 0x04F
HWPTAG_SHAPE_COMPONENT_ELLIPSE= HWPTAG_BEGIN + 64  # 0x050
HWPTAG_SHAPE_COMPONENT_ARC    = HWPTAG_BEGIN + 65  # 0x051
HWPTAG_SHAPE_COMPONENT_POLYGON= HWPTAG_BEGIN + 66  # 0x052
HWPTAG_SHAPE_COMPONENT_CURVE  = HWPTAG_BEGIN + 67  # 0x053
HWPTAG_SHAPE_COMPONENT_OLE    = HWPTAG_BEGIN + 68  # 0x054
HWPTAG_SHAPE_COMPONENT_PICTURE= HWPTAG_BEGIN + 69  # 0x055
HWPTAG_SHAPE_COMPONENT_CONTAINER = HWPTAG_BEGIN + 70  # 0x056
HWPTAG_CTRL_DATA              = HWPTAG_BEGIN + 71  # 0x057
HWPTAG_EQEDIT                 = HWPTAG_BEGIN + 72  # 0x058
HWPTAG_SHAPE_COMPONENT_TEXTART= HWPTAG_BEGIN + 74  # 0x05A
HWPTAG_FORM_OBJECT            = HWPTAG_BEGIN + 75  # 0x05B
HWPTAG_MEMO_SHAPE             = HWPTAG_BEGIN + 76  # 0x05C
HWPTAG_MEMO_LIST              = HWPTAG_BEGIN + 77  # 0x05D
HWPTAG_CHART_DATA             = HWPTAG_BEGIN + 79  # 0x05F
HWPTAG_VIDEO_DATA             = HWPTAG_BEGIN + 82  # 0x062
```

---

## 5. 제어 문자 (Control Characters)

텍스트 내 특수 문자 (WCHAR 값)

```python
CTRL_CHAR = {
    0:  'NULL',           # unusable
    1:  'RESERVED',
    2:  'SECTION_COLUMN', # 구역/단 정의 (extended, 8 chars)
    3:  'FIELD_START',    # 필드 시작 (extended, 8 chars)
    4:  'FIELD_END',      # 필드 끝 (inline, 8 chars)
    5:  'RESERVED',
    6:  'RESERVED',
    7:  'RESERVED',
    8:  'RESERVED',
    9:  'TAB',            # 탭 (inline, 8 chars)
    10: 'LINE_BREAK',     # 줄바꿈 (char, 1 char)
    11: 'DRAWING_TABLE',  # 그리기개체/표 (extended, 8 chars)
    12: 'RESERVED',
    13: 'PARA_BREAK',     # 문단 끝 (char, 1 char)
    14: 'RESERVED',
    15: 'HIDDEN_COMMENT', # 숨은 설명 (extended, 8 chars)
    16: 'HEADER_FOOTER',  # 머리말/꼬리말 (extended, 8 chars)
    17: 'FOOTNOTE_ENDNOTE', # 각주/미주 (extended, 8 chars)
    18: 'AUTO_NUMBER',    # 자동 번호 (extended, 8 chars)
    19: 'RESERVED',
    20: 'RESERVED',
    21: 'PAGE_CTRL',      # 페이지 컨트롤 (extended, 8 chars)
    22: 'BOOKMARK_INDEX', # 책갈피/찾아보기 (extended, 8 chars)
    23: 'OVERLAY_CHAR',   # 덧말/글자겹침 (extended, 8 chars)
    24: 'HYPHEN',         # 하이픈 (char, 1 char)
    25: 'RESERVED',
    26: 'RESERVED',
    27: 'RESERVED',
    28: 'RESERVED',
    29: 'RESERVED',
    30: 'NBSP',           # 묶음 빈칸 (char, 1 char)
    31: 'FIXED_WIDTH_SPACE', # 고정폭 빈칸 (char, 1 char)
}

# Extended 컨트롤: 8개 WCHAR (16 bytes) 차지
# - [0]: 컨트롤 코드
# - [1-4]: 추가 정보 (컨트롤 종류에 따라 다름)
# - [5-7]: reserved

# Inline 컨트롤: 8개 WCHAR (16 bytes) 차지, 그러나 텍스트 위치 1개만 차지
```

---

## 6. 컨트롤 ID (4-byte)

```python
def make_4chid(a, b, c, d):
    return (ord(a) << 24) | (ord(b) << 16) | (ord(c) << 8) | ord(d)

# 개체 컨트롤 (그리기/표/그림 등)
CTRL_ID = {
    'TABLE':     make_4chid('t', 'b', 'l', ' '),  # 0x74626C20
    'LINE':      make_4chid('$', 'l', 'i', 'n'),  # 직선
    'RECT':      make_4chid('$', 'r', 'e', 'c'),  # 사각형
    'ELLIPSE':   make_4chid('$', 'e', 'l', 'l'),  # 타원
    'ARC':       make_4chid('$', 'a', 'r', 'c'),  # 호
    'POLYGON':   make_4chid('$', 'p', 'o', 'l'),  # 다각형
    'CURVE':     make_4chid('$', 'c', 'u', 'r'),  # 곡선
    'EQUATION':  make_4chid('e', 'q', 'e', 'd'),  # 수식
    'PICTURE':   make_4chid('$', 'p', 'i', 'c'),  # 그림
    'OLE':       make_4chid('$', 'o', 'l', 'e'),  # OLE 개체
    'CONTAINER': make_4chid('$', 'c', 'o', 'n'),  # 묶음 개체
    'TEXTART':   make_4chid('$', 't', 'x', 't'),  # 글맵시
    
    # 비개체 컨트롤
    'SECD':      make_4chid('s', 'e', 'c', 'd'),  # 구역 정의
    'COLD':      make_4chid('c', 'o', 'l', 'd'),  # 단 정의
    'HEAD':      make_4chid('h', 'e', 'a', 'd'),  # 머리말
    'FOOT':      make_4chid('f', 'o', 'o', 't'),  # 꼬리말
    'FN':        make_4chid('f', 'n', ' ', ' '),  # 각주
    'EN':        make_4chid('e', 'n', ' ', ' '),  # 미주
    'ATNO':      make_4chid('a', 't', 'n', 'o'),  # 자동번호
    'NWNO':      make_4chid('n', 'w', 'n', 'o'),  # 새 번호
    'PGHD':      make_4chid('p', 'g', 'h', 'd'),  # 감추기
    'PGCT':      make_4chid('p', 'g', 'c', 't'),  # 홀/짝 조정
    'PGNP':      make_4chid('p', 'g', 'n', 'p'),  # 쪽 번호 위치
    'IDXM':      make_4chid('i', 'd', 'x', 'm'),  # 찾아보기 표식
    'BOKM':      make_4chid('b', 'o', 'k', 'm'),  # 책갈피
    'TCPS':      make_4chid('t', 'c', 'p', 's'),  # 글자 겹침
    'TDUT':      make_4chid('t', 'd', 'u', 't'),  # 덧말
    'TCMT':      make_4chid('t', 'c', 'm', 't'),  # 숨은 설명
}

# 필드 컨트롤 ID
FIELD_ID = {
    'UNKNOWN':    make_4chid('%', 'u', 'n', 'k'),
    'DATE':       make_4chid('%', 'd', 't', 'e'),
    'DOCDATE':    make_4chid('%', 'd', 'd', 't'),
    'PATH':       make_4chid('%', 'p', 'a', 't'),
    'BOOKMARK':   make_4chid('%', 'b', 'm', 'k'),
    'MAILMERGE':  make_4chid('%', 'm', 'm', 'g'),
    'CROSSREF':   make_4chid('%', 'x', 'r', 'f'),
    'FORMULA':    make_4chid('%', 'f', 'm', 'u'),
    'CLICKHERE':  make_4chid('%', 'c', 'l', 'k'),
    'SUMMARY':    make_4chid('%', 's', 'm', 'r'),
    'USERINFO':   make_4chid('%', 'u', 's', 'r'),
    'HYPERLINK':  make_4chid('%', 'h', 'l', 'k'),
    'MEMO':       make_4chid('%', '%', 'm', 'e'),
    'TOC':        make_4chid('%', 't', 'o', 'c'),
}
```

---

## 7. 핵심 레코드 파싱

### 7.1 HWPTAG_PARA_HEADER (문단 헤더)

```python
def parse_para_header(data: bytes) -> dict:
    nchars = struct.unpack('<I', data[0:4])[0]
    # 최상위 비트가 1이면 마스킹
    if nchars & 0x80000000:
        nchars &= 0x7FFFFFFF
    
    ctrl_mask = struct.unpack('<I', data[4:8])[0]
    para_shape_id = struct.unpack('<H', data[8:10])[0]
    style_id = data[10]
    col_split = data[11]  # 단 나누기 종류
    char_shape_count = struct.unpack('<H', data[12:14])[0]
    range_tag_count = struct.unpack('<H', data[14:16])[0]
    line_seg_count = struct.unpack('<H', data[16:18])[0]
    para_instance_id = struct.unpack('<I', data[18:22])[0]
    
    return {
        'nchars': nchars,
        'ctrl_mask': ctrl_mask,
        'para_shape_id': para_shape_id,
        'style_id': style_id,
        'col_split': col_split,
        'char_shape_count': char_shape_count,
        'range_tag_count': range_tag_count,
        'line_seg_count': line_seg_count,
        'para_instance_id': para_instance_id,
    }
```

### 7.2 HWPTAG_PARA_TEXT (문단 텍스트)

```python
def parse_para_text(data: bytes) -> str:
    """UTF-16LE 텍스트 파싱, 컨트롤 문자 처리"""
    text_parts = []
    pos = 0
    while pos < len(data):
        char_code = struct.unpack('<H', data[pos:pos+2])[0]
        
        if char_code < 32:
            # 제어 문자 처리
            if char_code in (0, 10, 13, 24, 30, 31):
                # char 타입: 1개 WCHAR
                if char_code == 10:
                    text_parts.append('\n')  # line break
                elif char_code == 13:
                    pass  # para break (문단 끝)
                elif char_code == 30:
                    text_parts.append('\u00A0')  # NBSP
                elif char_code == 31:
                    text_parts.append(' ')  # fixed width space
                pos += 2
            elif char_code in (4, 9):
                # inline 타입: 8개 WCHAR (16 bytes)
                if char_code == 9:
                    text_parts.append('\t')  # tab
                pos += 16
            else:
                # extended 타입: 8개 WCHAR (16 bytes)
                # 2, 3, 11, 15, 16, 17, 18, 21, 22, 23
                pos += 16
        else:
            # 일반 문자
            text_parts.append(chr(char_code))
            pos += 2
    
    return ''.join(text_parts)
```

### 7.3 HWPTAG_PARA_CHAR_SHAPE (문단 글자 모양)

```python
def parse_para_char_shape(data: bytes, count: int) -> list:
    """글자 모양 변경 위치와 ID 매핑"""
    shapes = []
    for i in range(count):
        offset = i * 8
        pos = struct.unpack('<I', data[offset:offset+4])[0]
        shape_id = struct.unpack('<I', data[offset+4:offset+8])[0]
        shapes.append({'pos': pos, 'shape_id': shape_id})
    return shapes
```

### 7.4 HWPTAG_CHAR_SHAPE (글자 모양 정의 - DocInfo)

```python
def parse_char_shape(data: bytes) -> dict:
    # 언어별 글꼴 ID (한글, 영문, 한자, 일어, 기타, 기호, 사용자)
    font_ids = struct.unpack('<7H', data[0:14])
    
    # 언어별 장평 (%)
    ratios = struct.unpack('<7B', data[14:21])
    
    # 언어별 자간 (%)  
    spacings = struct.unpack('<7b', data[21:28])
    
    # 언어별 상대 크기 (%)
    rel_sizes = struct.unpack('<7B', data[28:35])
    
    # 언어별 글자 위치 (%)
    positions = struct.unpack('<7b', data[35:42])
    
    # 기준 크기 (HWPUNIT, 1/7200 inch)
    base_size = struct.unpack('<I', data[42:46])[0]
    # 실제 pt = base_size / 100
    
    # 속성
    attr = struct.unpack('<I', data[46:50])[0]
    
    # 그림자 간격
    shadow_gap = struct.unpack('<h', data[50:52])[0]
    
    # 색상들 (COLORREF: 0x00BBGGRR)
    char_color = struct.unpack('<I', data[52:56])[0]
    underline_color = struct.unpack('<I', data[56:60])[0]
    shade_color = struct.unpack('<I', data[60:64])[0]
    shadow_color = struct.unpack('<I', data[64:68])[0]
    
    return {
        'font_ids': font_ids,
        'base_size_pt': base_size / 100,
        'bold': bool(attr & 0x02),
        'italic': bool(attr & 0x01),
        'underline': (attr >> 2) & 0x03,
        'strikeout': (attr >> 18) & 0x07,
        'superscript': bool(attr & 0x8000),
        'subscript': bool(attr & 0x10000),
        'char_color': char_color,
    }
```

### 7.5 HWPTAG_PARA_SHAPE (문단 모양 정의 - DocInfo)

```python
def parse_para_shape(data: bytes) -> dict:
    attr1 = struct.unpack('<I', data[0:4])[0]
    left_margin = struct.unpack('<i', data[4:8])[0]    # HWPUNIT
    right_margin = struct.unpack('<i', data[8:12])[0]
    indent = struct.unpack('<i', data[12:16])[0]       # 들여쓰기(양수)/내어쓰기(음수)
    prev_spacing = struct.unpack('<i', data[16:20])[0] # 문단 위 간격
    next_spacing = struct.unpack('<i', data[20:24])[0] # 문단 아래 간격
    line_spacing = struct.unpack('<i', data[24:28])[0] # 줄간격 (구버전)
    tab_def_id = struct.unpack('<H', data[28:30])[0]
    numbering_id = struct.unpack('<H', data[30:32])[0]
    border_fill_id = struct.unpack('<H', data[32:34])[0]
    
    # 정렬: bit 0-1
    alignment = attr1 & 0x03
    # 0: 양쪽, 1: 왼쪽, 2: 오른쪽, 3: 가운데, 4: 배분, 5: 나눔
    
    # 문단 머리 모양: bit 23-24
    head_type = (attr1 >> 23) & 0x03
    # 0: 없음, 1: 개요, 2: 번호, 3: 글머리표
    
    # 문단 수준: bit 25-27
    level = (attr1 >> 25) & 0x07
    
    return {
        'alignment': alignment,
        'left_margin': left_margin,
        'right_margin': right_margin,
        'indent': indent,
        'prev_spacing': prev_spacing,
        'next_spacing': next_spacing,
        'line_spacing': line_spacing,
        'head_type': head_type,  # 번호/글머리표
        'level': level,
        'numbering_id': numbering_id,
    }
```

### 7.6 HWPTAG_CTRL_HEADER (컨트롤 헤더)

```python
def parse_ctrl_header(data: bytes) -> dict:
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
        'ctrl_ch': ctrl_ch,  # 예: 'tbl ', '$pic', 'secd' 등
        'data': data[4:],
    }
```

### 7.7 HWPTAG_TABLE (표)

```python
def parse_table(data: bytes) -> dict:
    attr = struct.unpack('<I', data[0:4])[0]
    row_count = struct.unpack('<H', data[4:6])[0]
    col_count = struct.unpack('<H', data[6:8])[0]
    cell_spacing = struct.unpack('<H', data[8:10])[0]
    
    # 안쪽 여백 (왼, 오른, 위, 아래)
    margins = struct.unpack('<4H', data[10:18])
    
    # 각 행의 높이 (row_count 개)
    row_heights = []
    offset = 18
    for _ in range(row_count):
        h = struct.unpack('<H', data[offset:offset+2])[0]
        row_heights.append(h)
        offset += 2
    
    # 테두리/배경 ID
    border_fill_id = struct.unpack('<H', data[offset:offset+2])[0]
    
    return {
        'row_count': row_count,
        'col_count': col_count,
        'cell_spacing': cell_spacing,
        'margins': margins,
        'row_heights': row_heights,
        'border_fill_id': border_fill_id,
    }
```

### 7.8 HWPTAG_LIST_HEADER (문단 리스트 헤더 - 셀 등)

```python
def parse_list_header(data: bytes) -> dict:
    para_count = struct.unpack('<h', data[0:2])[0]
    attr = struct.unpack('<I', data[2:6])[0]
    
    # 텍스트 방향: bit 0-2
    text_direction = attr & 0x07  # 0: 가로, 1: 세로
    
    # 세로 정렬: bit 5-6
    vert_align = (attr >> 5) & 0x03  # 0: top, 1: center, 2: bottom
    
    return {
        'para_count': para_count,
        'text_direction': text_direction,
        'vert_align': vert_align,
    }
```

---

## 8. DocInfo ID 매핑

DocInfo의 HWPTAG_ID_MAPPINGS는 각 속성 타입별 개수를 정의:

```python
def parse_id_mappings(data: bytes) -> dict:
    """
    순서대로:
    바이너리, 한글폰트, 영문폰트, 한자폰트, 일어폰트,
    기타폰트, 기호폰트, 사용자폰트, 테두리/배경,
    글자모양, 탭, 번호, 글머리표, 문단모양, 스타일, 메모
    """
    counts = struct.unpack('<32i', data[:128])
    return {
        'bin_data': counts[0],
        'font_hangul': counts[1],
        'font_english': counts[2],
        'font_hanja': counts[3],
        'font_japanese': counts[4],
        'font_other': counts[5],
        'font_symbol': counts[6],
        'font_user': counts[7],
        'border_fill': counts[8],
        'char_shape': counts[9],
        'tab_def': counts[10],
        'numbering': counts[11],
        'bullet': counts[12],
        'para_shape': counts[13],
        'style': counts[14],
        'memo': counts[15],
    }
```

---

## 9. 바이너리 데이터 (이미지 등)

### 9.1 HWPTAG_BIN_DATA (DocInfo)

```python
def parse_bin_data_info(data: bytes) -> dict:
    attr = struct.unpack('<H', data[0:2])[0]
    
    # 타입: bit 0-3
    bin_type = attr & 0x0F
    # 0: LINK (외부 파일)
    # 1: EMBEDDING (내장)
    # 2: STORAGE (OLE)
    
    # 압축: bit 4-5
    compress = (attr >> 4) & 0x03
    # 0: 기본 (적용 안함)
    # 1: 압축
    # 2: 비압축
    
    # 상태: bit 8-9
    state = (attr >> 8) & 0x03
    # 0: 아직 접근 안함
    # 1: 접근 성공
    # 2: 접근 오류
    # 3: 무시
    
    if bin_type == 0:  # LINK
        # 절대/상대 경로
        path_len = struct.unpack('<H', data[2:4])[0]
        path = data[4:4+path_len*2].decode('utf-16-le')
        return {'type': 'link', 'path': path}
    else:  # EMBEDDING or STORAGE
        bin_id = struct.unpack('<H', data[2:4])[0]
        ext_len = struct.unpack('<H', data[4:6])[0]
        ext = data[6:6+ext_len*2].decode('utf-16-le')
        return {'type': 'embed', 'bin_id': bin_id, 'extension': ext}
```

### 9.2 BinData 스트림에서 이미지 추출

```python
def extract_bindata(ole, bin_id: int, compressed: bool) -> bytes:
    """BinData 스토리지에서 바이너리 추출"""
    # 스트림 이름: BIN%04X.확장자 또는 DATA%d
    stream_name = f'BinData/BIN{bin_id:04X}.jpg'  # 확장자는 실제로 확인 필요
    
    try:
        stream = ole.openstream(stream_name)
        data = stream.read()
        
        if compressed:
            data = zlib.decompress(data, -15)  # raw deflate
        
        return data
    except:
        return None
```

---

## 10. 하이퍼링크 파싱

```python
def parse_hyperlink_ctrl(data: bytes) -> dict:
    """FIELD_HYPERLINK 컨트롤 데이터"""
    # Parameter Set 구조
    # 일반적으로 URL은 파라미터 아이템에 문자열로 저장
    
    # 간략화된 파싱 (실제로는 Parameter Set 구조 따라야 함)
    # 보통 URL 문자열을 찾아서 추출
    try:
        # UTF-16LE에서 URL 찾기
        text = data.decode('utf-16-le', errors='ignore')
        # http:// 또는 https:// 찾기
        import re
        urls = re.findall(r'https?://[^\x00]+', text)
        if urls:
            return {'url': urls[0].split('\x00')[0]}
    except:
        pass
    return {}
```

---

## 11. 마크다운 변환 핵심 로직

```python
class HWPToMarkdown:
    def __init__(self):
        self.char_shapes = []   # DocInfo에서 로드
        self.para_shapes = []
        self.styles = []
        self.numbering = []
        
    def convert_paragraph(self, para_header, para_text, char_shape_info) -> str:
        """문단을 마크다운으로 변환"""
        text = self.parse_para_text_to_str(para_text)
        para_shape = self.para_shapes[para_header['para_shape_id']]
        
        # 문단 모양에 따른 마크다운 변환
        md_lines = []
        
        # 번호/글머리표 처리
        if para_shape['head_type'] == 2:  # 번호
            level = para_shape['level']
            prefix = '   ' * level + '1. '
            text = prefix + text
        elif para_shape['head_type'] == 3:  # 글머리표
            level = para_shape['level']
            prefix = '   ' * level + '- '
            text = prefix + text
        
        # 글자 모양 적용 (굵게, 기울임 등)
        text = self.apply_char_styles(text, char_shape_info)
        
        return text
    
    def apply_char_styles(self, text: str, shapes: list) -> str:
        """글자 모양을 마크다운 인라인 스타일로 변환"""
        if not shapes:
            return text
        
        # 역순으로 처리 (위치가 뒤바뀌지 않도록)
        result = list(text)
        for i in range(len(shapes) - 1, -1, -1):
            shape = shapes[i]
            char_shape = self.char_shapes[shape['shape_id']]
            start = shape['pos']
            end = shapes[i+1]['pos'] if i+1 < len(shapes) else len(text)
            
            segment = text[start:end]
            
            if char_shape['bold'] and char_shape['italic']:
                segment = f'***{segment}***'
            elif char_shape['bold']:
                segment = f'**{segment}**'
            elif char_shape['italic']:
                segment = f'*{segment}*'
            
            if char_shape['strikeout']:
                segment = f'~~{segment}~~'
            
            result[start:end] = list(segment)
        
        return ''.join(result)
    
    def convert_table(self, table_info, cells) -> str:
        """표를 마크다운 테이블로 변환"""
        rows = table_info['row_count']
        cols = table_info['col_count']
        
        md_lines = []
        cell_idx = 0
        
        for r in range(rows):
            row_cells = []
            for c in range(cols):
                cell_text = cells[cell_idx] if cell_idx < len(cells) else ''
                row_cells.append(cell_text.replace('\n', ' ').strip())
                cell_idx += 1
            
            md_lines.append('| ' + ' | '.join(row_cells) + ' |')
            
            # 첫 행 후 구분선
            if r == 0:
                md_lines.append('|' + '---|' * cols)
        
        return '\n'.join(md_lines)
```

---

## 12. 전체 파싱 흐름

```python
import olefile
import zlib
import struct

def parse_hwp(filepath: str) -> dict:
    ole = olefile.OleFileIO(filepath)
    
    # 1. FileHeader 파싱
    header_data = ole.openstream('FileHeader').read()
    header = parse_file_header(header_data)
    
    if header['encrypted']:
        raise ValueError("암호화된 문서는 지원하지 않습니다")
    
    # 2. DocInfo 파싱
    docinfo_data = ole.openstream('DocInfo').read()
    if header['compressed']:
        docinfo_data = zlib.decompress(docinfo_data, -15)
    
    docinfo_records = parse_records(docinfo_data)
    
    # ID 매핑, 글자모양, 문단모양 등 수집
    doc_info = process_docinfo_records(docinfo_records)
    
    # 3. BodyText 파싱
    sections = []
    section_idx = 0
    while True:
        stream_name = f'BodyText/Section{section_idx}'
        if not ole.exists(stream_name):
            break
        
        section_data = ole.openstream(stream_name).read()
        if header['compressed']:
            section_data = zlib.decompress(section_data, -15)
        
        section_records = parse_records(section_data)
        sections.append(section_records)
        section_idx += 1
    
    # 4. BinData (이미지) 수집
    bindata = {}
    if ole.exists('BinData'):
        for entry in ole.listdir():
            if entry[0] == 'BinData':
                name = '/'.join(entry)
                data = ole.openstream(name).read()
                if header['compressed']:
                    try:
                        data = zlib.decompress(data, -15)
                    except:
                        pass
                bindata[entry[-1]] = data
    
    ole.close()
    
    return {
        'header': header,
        'doc_info': doc_info,
        'sections': sections,
        'bindata': bindata,
    }
```

---

## 13. 주의 사항

1. **zlib 압축**: `decompress(data, -15)` - raw deflate (no header)
2. **UTF-16LE**: 모든 문자열은 little-endian UTF-16
3. **HWPUNIT**: 좌표/크기 단위는 1/7200 inch
4. **레코드 Level**: 계층 구조 (표 > 셀 > 문단)
5. **컨트롤 문자 크기**:
   - char 타입 (0,10,13,24,30,31): 1 WCHAR
   - inline 타입 (4,9): 8 WCHAR
   - extended 타입 (나머지): 8 WCHAR
6. **배포용 문서**: `distributed=True`면 추가 암호화 레이어 있음

---

## 부록: 테두리선/색상 참조

### 테두리선 종류

| 값 | 설명 |
|----|------|
| 0 | 실선 |
| 1 | 긴 점선 |
| 2 | 점선 |
| 3 | -.-.-.  |
| 4 | -..-..- |
| 5 | 긴 대시 |
| 6 | 큰 원 |
| 7 | 이중선 |
| 8 | 가는선+굵은선 |
| 9 | 굵은선+가는선 |
| 10 | 삼중선 |
| 11 | 물결 |
| 12 | 이중 물결 |

### COLORREF

```python
def colorref_to_rgb(color: int) -> tuple:
    """0x00BBGGRR → (R, G, B)"""
    r = color & 0xFF
    g = (color >> 8) & 0xFF
    b = (color >> 16) & 0xFF
    return (r, g, b)

def colorref_to_hex(color: int) -> str:
    """0x00BBGGRR → #RRGGBB"""
    r, g, b = colorref_to_rgb(color)
    return f'#{r:02X}{g:02X}{b:02X}'
```
