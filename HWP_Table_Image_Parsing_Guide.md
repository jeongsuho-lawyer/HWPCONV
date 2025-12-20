# HWP 표/이미지 파싱 완벽 가이드

## 1. 레코드 계층 구조 이해

표와 이미지는 **레코드의 Level(계층)**을 통해 부모-자식 관계를 형성합니다.

```
[CTRL_HEADER level=0] ─── 컨트롤 시작 (ctrl_id로 표/그림 구분)
  ├── [TABLE level=1] ─── 표 속성 (행/열 수, 여백 등)
  │     ├── [LIST_HEADER level=2] ─── 셀 1 리스트 헤더
  │     │     └── [PARA_HEADER level=3] ─── 셀 1 문단
  │     │           ├── [PARA_TEXT level=4]
  │     │           └── [PARA_CHAR_SHAPE level=4]
  │     ├── [LIST_HEADER level=2] ─── 셀 2 리스트 헤더
  │     │     └── [PARA_HEADER level=3] ─── 셀 2 문단
  │     └── ...
  └── [SHAPE_COMPONENT level=1] ─── (그림일 경우)
```

**핵심**: `level`이 감소하면 현재 개체 종료, 같거나 증가하면 하위 요소.

---

## 2. 표 파싱 완벽 가이드

### 2.1 표 식별

```python
# CTRL_HEADER에서 ctrl_id 확인
CTRL_ID_TABLE = 0x74626C20  # 'tbl ' (big-endian)

def is_table_control(ctrl_header_data: bytes) -> bool:
    ctrl_id = struct.unpack('<I', ctrl_header_data[0:4])[0]
    return ctrl_id == CTRL_ID_TABLE
```

### 2.2 개체 공통 속성 (모든 개체 컨트롤 공통)

표, 그림, OLE 등 모든 개체 컨트롤은 먼저 **개체 공통 속성**을 파싱해야 합니다.

```python
def parse_common_object_attrs(data: bytes, offset: int = 0) -> tuple[dict, int]:
    """
    개체 공통 속성 파싱
    Returns: (속성 dict, 다음 offset)
    """
    pos = offset
    
    ctrl_id = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    attr = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    vert_offset = struct.unpack('<i', data[pos:pos+4])[0]  # HWPUNIT (signed)
    pos += 4
    
    horz_offset = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    width = struct.unpack('<I', data[pos:pos+4])[0]  # HWPUNIT
    pos += 4
    
    height = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    z_order = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    # 바깥 여백 (왼, 오른, 위, 아래) - HWPUNIT16
    margins = struct.unpack('<4H', data[pos:pos+8])
    pos += 8
    
    instance_id = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    prevent_page_break = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    # 개체 설명문
    desc_len = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    description = ''
    if desc_len > 0:
        description = data[pos:pos+desc_len*2].decode('utf-16-le', errors='ignore')
        pos += desc_len * 2
    
    # 속성 비트 해석
    treat_as_char = bool(attr & 0x01)           # bit 0: 글자처럼 취급
    affect_line_spacing = bool(attr & 0x04)    # bit 2: 줄간격 영향
    vert_rel_to = (attr >> 3) & 0x03           # bit 3-4: 세로 기준
    vert_align = (attr >> 5) & 0x07            # bit 5-7: 세로 정렬
    horz_rel_to = (attr >> 8) & 0x03           # bit 8-9: 가로 기준
    horz_align = (attr >> 10) & 0x07           # bit 10-12: 가로 정렬
    flow_with_text = (attr >> 21) & 0x07       # bit 21-23: 텍스트 배치
    
    return {
        'ctrl_id': ctrl_id,
        'width': width,              # HWPUNIT (1/7200 inch)
        'height': height,
        'width_mm': width / 7200 * 25.4,
        'height_mm': height / 7200 * 25.4,
        'vert_offset': vert_offset,
        'horz_offset': horz_offset,
        'z_order': z_order,
        'margins': margins,
        'instance_id': instance_id,
        'prevent_page_break': prevent_page_break,
        'description': description,
        'treat_as_char': treat_as_char,
        'vert_rel_to': vert_rel_to,  # 0:paper, 1:page, 2:para
        'horz_rel_to': horz_rel_to,  # 0:page, 1:page, 2:column, 3:para
        'flow_with_text': flow_with_text,  # 0:square, 1:tight, 2:through, 3:top/bottom, 4:behind, 5:front
    }, pos
```

### 2.3 HWPTAG_TABLE 파싱

```python
HWPTAG_TABLE = 0x010 + 61  # 0x04D

def parse_table(data: bytes) -> dict:
    """
    표 개체 속성 파싱 (개체 공통 속성 이후 부분)
    """
    pos = 0
    
    # 속성 (4 bytes)
    attr = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    # 행/열 수
    row_count = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    col_count = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    # 셀 간격 (HWPUNIT16)
    cell_spacing = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    # 안쪽 여백 (왼, 오른, 위, 아래) - 각 2 bytes
    left_margin = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    right_margin = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    top_margin = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    bottom_margin = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    # 각 행의 높이 (row_count 개)
    row_heights = []
    for _ in range(row_count):
        h = struct.unpack('<H', data[pos:pos+2])[0]
        row_heights.append(h)
        pos += 2
    
    # 테두리/배경 ID
    border_fill_id = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    # 5.0.1.0 이상: Valid Zone Info
    valid_zone_count = 0
    zones = []
    if pos + 2 <= len(data):
        valid_zone_count = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2
        
        for _ in range(valid_zone_count):
            if pos + 10 <= len(data):
                zone = {
                    'start_col': struct.unpack('<H', data[pos:pos+2])[0],
                    'start_row': struct.unpack('<H', data[pos+2:pos+4])[0],
                    'end_col': struct.unpack('<H', data[pos+4:pos+6])[0],
                    'end_row': struct.unpack('<H', data[pos+6:pos+8])[0],
                    'border_fill_id': struct.unpack('<H', data[pos+8:pos+10])[0],
                }
                zones.append(zone)
                pos += 10
    
    # 속성 비트 해석
    page_break = attr & 0x03           # bit 0-1: 쪽 경계 나눔 (0:안나눔, 1:셀단위, 2:안나눔)
    repeat_header = bool(attr & 0x04)  # bit 2: 제목줄 자동 반복
    
    return {
        'row_count': row_count,
        'col_count': col_count,
        'cell_spacing': cell_spacing,
        'margins': {
            'left': left_margin,
            'right': right_margin,
            'top': top_margin,
            'bottom': bottom_margin,
        },
        'row_heights': row_heights,
        'border_fill_id': border_fill_id,
        'page_break': page_break,
        'repeat_header': repeat_header,
        'zones': zones,
    }
```

### 2.4 셀 파싱 (LIST_HEADER + 셀 속성)

```python
HWPTAG_LIST_HEADER = 0x010 + 56  # 0x048

def parse_list_header(data: bytes) -> dict:
    """
    문단 리스트 헤더 (셀, 캡션 등에서 사용)
    """
    pos = 0
    
    para_count = struct.unpack('<h', data[pos:pos+2])[0]  # INT16
    pos += 2
    
    attr = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    # 속성 비트 해석
    text_direction = attr & 0x07        # bit 0-2: 0=가로, 1=세로
    line_wrap = (attr >> 3) & 0x03      # bit 3-4: 줄바꿈 (0:일반, 1:자간조절, 2:폭늘림)
    vert_align = (attr >> 5) & 0x03     # bit 5-6: 세로정렬 (0:top, 1:center, 2:bottom)
    
    return {
        'para_count': para_count,
        'text_direction': text_direction,
        'line_wrap': line_wrap,
        'vert_align': vert_align,
        'raw_data': data[pos:],  # 나머지는 셀 속성
    }

def parse_cell_property(data: bytes) -> dict:
    """
    셀 속성 파싱 (LIST_HEADER 다음 26 bytes)
    """
    pos = 0
    
    col_addr = struct.unpack('<H', data[pos:pos+2])[0]  # 0부터 시작
    pos += 2
    
    row_addr = struct.unpack('<H', data[pos:pos+2])[0]  # 0부터 시작
    pos += 2
    
    col_span = struct.unpack('<H', data[pos:pos+2])[0]  # 열 병합 수
    pos += 2
    
    row_span = struct.unpack('<H', data[pos:pos+2])[0]  # 행 병합 수
    pos += 2
    
    width = struct.unpack('<I', data[pos:pos+4])[0]    # HWPUNIT
    pos += 4
    
    height = struct.unpack('<I', data[pos:pos+4])[0]   # HWPUNIT
    pos += 4
    
    # 셀 4방향 여백 (왼, 오른, 위, 아래)
    margins = struct.unpack('<4H', data[pos:pos+8])
    pos += 8
    
    border_fill_id = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    return {
        'col': col_addr,
        'row': row_addr,
        'col_span': col_span,
        'row_span': row_span,
        'width': width,
        'height': height,
        'width_mm': width / 7200 * 25.4,
        'height_mm': height / 7200 * 25.4,
        'margins': {
            'left': margins[0],
            'right': margins[1],
            'top': margins[2],
            'bottom': margins[3],
        },
        'border_fill_id': border_fill_id,
    }
```

### 2.5 전체 표 파싱 흐름

```python
def parse_table_from_records(records: list, start_idx: int) -> dict:
    """
    레코드 리스트에서 표 전체 파싱
    
    records: 전체 레코드 리스트
    start_idx: CTRL_HEADER 레코드 인덱스
    """
    table_info = None
    cells = []
    current_cell = None
    base_level = records[start_idx]['level']
    
    i = start_idx + 1
    while i < len(records):
        rec = records[i]
        tag_id = rec['tag_id']
        level = rec['level']
        data = rec['data']
        
        # 레벨이 base_level 이하로 내려가면 표 종료
        if level <= base_level:
            break
        
        if tag_id == HWPTAG_TABLE:
            table_info = parse_table(data)
            
        elif tag_id == HWPTAG_LIST_HEADER:
            # 이전 셀 저장
            if current_cell is not None:
                cells.append(current_cell)
            
            list_header = parse_list_header(data)
            
            # LIST_HEADER 이후 26 bytes가 셀 속성
            # 실제로는 LIST_HEADER 레코드 내에 포함되어 있음
            if len(list_header['raw_data']) >= 26:
                cell_prop = parse_cell_property(list_header['raw_data'])
                current_cell = {
                    'list_header': list_header,
                    'cell_property': cell_prop,
                    'paragraphs': [],
                }
            else:
                current_cell = {
                    'list_header': list_header,
                    'cell_property': None,
                    'paragraphs': [],
                }
                
        elif tag_id == HWPTAG_PARA_HEADER:
            # 현재 셀에 문단 추가
            if current_cell is not None:
                current_cell['paragraphs'].append({
                    'header': parse_para_header(data),
                    'text': '',
                    'char_shapes': [],
                })
                
        elif tag_id == HWPTAG_PARA_TEXT:
            if current_cell and current_cell['paragraphs']:
                text = parse_para_text(data)
                current_cell['paragraphs'][-1]['text'] = text
                
        elif tag_id == HWPTAG_PARA_CHAR_SHAPE:
            if current_cell and current_cell['paragraphs']:
                count = current_cell['paragraphs'][-1]['header'].get('char_shape_count', 0)
                shapes = parse_para_char_shape(data, count)
                current_cell['paragraphs'][-1]['char_shapes'] = shapes
        
        i += 1
    
    # 마지막 셀 저장
    if current_cell is not None:
        cells.append(current_cell)
    
    return {
        'table_info': table_info,
        'cells': cells,
        'end_index': i,
    }
```

### 2.6 표 → 마크다운 변환

```python
def table_to_markdown(table_data: dict) -> str:
    """
    파싱된 표 데이터를 마크다운 테이블로 변환
    """
    table_info = table_data['table_info']
    cells = table_data['cells']
    
    if not table_info:
        return ''
    
    rows = table_info['row_count']
    cols = table_info['col_count']
    
    # 2D 그리드 생성 (병합 처리용)
    grid = [['' for _ in range(cols)] for _ in range(rows)]
    occupied = [[False for _ in range(cols)] for _ in range(rows)]
    
    for cell in cells:
        cell_prop = cell.get('cell_property')
        if not cell_prop:
            continue
        
        row = cell_prop['row']
        col = cell_prop['col']
        row_span = cell_prop['row_span']
        col_span = cell_prop['col_span']
        
        # 셀 텍스트 추출
        text_parts = []
        for para in cell.get('paragraphs', []):
            text = para.get('text', '').strip()
            if text:
                text_parts.append(text)
        cell_text = ' '.join(text_parts)
        
        # 마크다운 테이블에서는 줄바꿈을 <br>로 대체
        cell_text = cell_text.replace('\n', '<br>')
        # 파이프 문자 이스케이프
        cell_text = cell_text.replace('|', '\\|')
        
        # 그리드에 배치
        if row < rows and col < cols:
            grid[row][col] = cell_text
            
            # 병합 영역 표시
            for r in range(row, min(row + row_span, rows)):
                for c in range(col, min(col + col_span, cols)):
                    occupied[r][c] = True
    
    # 마크다운 테이블 생성
    md_lines = []
    
    for r in range(rows):
        row_cells = []
        for c in range(cols):
            row_cells.append(grid[r][c])
        md_lines.append('| ' + ' | '.join(row_cells) + ' |')
        
        # 첫 행 다음에 구분선
        if r == 0:
            md_lines.append('|' + '---|' * cols)
    
    return '\n'.join(md_lines)
```

---

## 3. 이미지(그림) 파싱 완벽 가이드

### 3.1 이미지 식별

```python
CTRL_ID_PICTURE = 0x24706963  # '$pic' (big-endian)

def is_picture_control(ctrl_header_data: bytes) -> bool:
    ctrl_id = struct.unpack('<I', ctrl_header_data[0:4])[0]
    return ctrl_id == CTRL_ID_PICTURE
```

### 3.2 DocInfo의 바이너리 데이터 정보 (HWPTAG_BIN_DATA)

먼저 DocInfo에서 바이너리 데이터 매핑 정보를 파싱해야 합니다.

```python
HWPTAG_BIN_DATA = 0x010 + 2  # 0x012

def parse_bin_data_info(data: bytes) -> dict:
    """
    DocInfo의 바이너리 데이터 정보 파싱
    """
    pos = 0
    
    attr = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    # 타입: bit 0-3
    bin_type = attr & 0x0F
    # 0: LINK (외부 파일 참조)
    # 1: EMBEDDING (파일 포함)
    # 2: STORAGE (OLE 저장소)
    
    # 압축: bit 4-5
    compress_mode = (attr >> 4) & 0x03
    # 0: 스토리지 기본값
    # 1: 압축
    # 2: 비압축
    
    # 상태: bit 8-9
    state = (attr >> 8) & 0x03
    # 0: 미접근
    # 1: 접근 성공
    # 2: 접근 실패
    # 3: 무시됨
    
    result = {
        'type': bin_type,
        'type_name': ['LINK', 'EMBEDDING', 'STORAGE'][bin_type] if bin_type < 3 else 'UNKNOWN',
        'compress_mode': compress_mode,
        'state': state,
    }
    
    if bin_type == 0:  # LINK
        # 절대 경로
        abs_path_len = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2
        abs_path = data[pos:pos+abs_path_len*2].decode('utf-16-le', errors='ignore')
        pos += abs_path_len * 2
        
        # 상대 경로
        rel_path_len = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2
        rel_path = data[pos:pos+rel_path_len*2].decode('utf-16-le', errors='ignore')
        pos += rel_path_len * 2
        
        result['abs_path'] = abs_path
        result['rel_path'] = rel_path
        
    else:  # EMBEDDING or STORAGE
        bin_id = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2
        
        ext_len = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2
        
        extension = ''
        if ext_len > 0:
            extension = data[pos:pos+ext_len*2].decode('utf-16-le', errors='ignore')
            pos += ext_len * 2
        
        result['bin_id'] = bin_id
        result['extension'] = extension
    
    return result
```

### 3.3 HWPTAG_SHAPE_COMPONENT_PICTURE 파싱

```python
HWPTAG_SHAPE_COMPONENT_PICTURE = 0x010 + 69  # 0x055

def parse_picture(data: bytes) -> dict:
    """
    그림 개체 속성 파싱
    개체 공통 속성 + 개체 요소 속성 이후 부분
    """
    pos = 0
    
    # 테두리 색 (COLORREF: 0x00BBGGRR)
    border_color = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    # 테두리 두께
    border_thickness = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    # 테두리 속성
    border_attr = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    # 이미지 테두리 사각형 좌표 (x 4개, y 4개)
    img_rect_x = struct.unpack('<4i', data[pos:pos+16])
    pos += 16
    img_rect_y = struct.unpack('<4i', data[pos:pos+16])
    pos += 16
    
    # 자르기 영역
    crop_left = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    crop_top = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    crop_right = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    crop_bottom = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    # 안쪽 여백 (8 bytes)
    inner_margins = struct.unpack('<4H', data[pos:pos+8])
    pos += 8
    
    # 그림 정보 (5 bytes) - 중요!
    # 여기서 bin_data_id를 얻음
    brightness = data[pos]
    pos += 1
    contrast = data[pos]
    pos += 1
    effect = data[pos]
    pos += 1
    bin_data_id = struct.unpack('<H', data[pos:pos+2])[0]  # ★ 핵심: BinData ID
    pos += 2
    
    # 테두리 투명도
    border_transparency = data[pos] if pos < len(data) else 0
    pos += 1
    
    # Instance ID
    instance_id = 0
    if pos + 4 <= len(data):
        instance_id = struct.unpack('<I', data[pos:pos+4])[0]
        pos += 4
    
    return {
        'border_color': border_color,
        'border_color_rgb': colorref_to_rgb(border_color),
        'border_thickness': border_thickness,
        'border_attr': border_attr,
        'img_rect': {
            'x': img_rect_x,
            'y': img_rect_y,
        },
        'crop': {
            'left': crop_left,
            'top': crop_top,
            'right': crop_right,
            'bottom': crop_bottom,
        },
        'inner_margins': inner_margins,
        'brightness': brightness,
        'contrast': contrast,
        'effect': effect,
        'bin_data_id': bin_data_id,  # ★ BinData 스토리지의 파일 ID
        'border_transparency': border_transparency,
        'instance_id': instance_id,
    }

def colorref_to_rgb(color: int) -> tuple:
    """COLORREF (0x00BBGGRR) → (R, G, B)"""
    r = color & 0xFF
    g = (color >> 8) & 0xFF
    b = (color >> 16) & 0xFF
    return (r, g, b)
```

### 3.4 개체 요소 속성 (SHAPE_COMPONENT) 파싱

그림 개체는 개체 요소 속성도 포함합니다.

```python
HWPTAG_SHAPE_COMPONENT = 0x010 + 60  # 0x04C

def parse_shape_component(data: bytes) -> dict:
    """
    개체 요소 속성 파싱
    그림, 그리기 개체 등에서 사용
    """
    pos = 0
    
    ctrl_id = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    # GenShapeObject인 경우 ctrl_id가 두 번 나옴
    # 두 번째도 ctrl_id일 수 있으니 확인
    
    x_offset_in_group = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    y_offset_in_group = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    group_level = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    local_version = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    initial_width = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    initial_height = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    current_width = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    current_height = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    attr = struct.unpack('<I', data[pos:pos+4])[0]
    pos += 4
    
    # 속성: bit 0 = horz flip, bit 1 = vert flip
    horz_flip = bool(attr & 0x01)
    vert_flip = bool(attr & 0x02)
    
    rotation_angle = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    
    rotation_center_x = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    rotation_center_y = struct.unpack('<i', data[pos:pos+4])[0]
    pos += 4
    
    # 렌더링 정보 (가변 길이)
    rendering_info = data[pos:]
    
    return {
        'ctrl_id': ctrl_id,
        'x_offset': x_offset_in_group,
        'y_offset': y_offset_in_group,
        'group_level': group_level,
        'initial_size': (initial_width, initial_height),
        'current_size': (current_width, current_height),
        'current_width_mm': current_width / 7200 * 25.4,
        'current_height_mm': current_height / 7200 * 25.4,
        'horz_flip': horz_flip,
        'vert_flip': vert_flip,
        'rotation_angle': rotation_angle,
        'rotation_center': (rotation_center_x, rotation_center_y),
    }
```

### 3.5 BinData 스토리지에서 이미지 추출

```python
import olefile
import zlib

def extract_image_from_hwp(ole: olefile.OleFileIO, 
                           bin_data_id: int, 
                           bin_data_info: dict,
                           compressed: bool) -> bytes:
    """
    HWP 파일에서 이미지 바이너리 추출
    
    ole: OLE 파일 객체
    bin_data_id: 그림 개체의 bin_data_id (1부터 시작)
    bin_data_info: DocInfo에서 파싱한 바이너리 정보
    compressed: FileHeader의 압축 플래그
    """
    
    # BinData 스토리지의 스트림 이름 규칙
    # 형식: BIN{ID:04X}.{extension}
    # 예: BIN0001.jpg, BIN0002.png
    
    extension = bin_data_info.get('extension', 'jpg').lower()
    stream_name = f'BinData/BIN{bin_data_id:04X}.{extension}'
    
    # 대안 이름들
    alt_names = [
        stream_name,
        f'BinData/BIN{bin_data_id:04X}.{extension.upper()}',
        f'BinData/DATA{bin_data_id}',
    ]
    
    image_data = None
    
    for name in alt_names:
        try:
            if ole.exists(name):
                stream = ole.openstream(name)
                image_data = stream.read()
                break
        except:
            continue
    
    if image_data is None:
        # 모든 BinData 스트림 나열해서 찾기
        for entry in ole.listdir():
            if entry[0] == 'BinData':
                entry_name = '/'.join(entry)
                # 파일명에서 ID 추출 시도
                filename = entry[-1]
                if filename.startswith('BIN'):
                    try:
                        file_id = int(filename[3:7], 16)
                        if file_id == bin_data_id:
                            image_data = ole.openstream(entry_name).read()
                            break
                    except:
                        pass
    
    if image_data is None:
        return None
    
    # 압축 해제
    compress_mode = bin_data_info.get('compress_mode', 0)
    
    need_decompress = False
    if compress_mode == 1:  # 무조건 압축
        need_decompress = True
    elif compress_mode == 0 and compressed:  # 기본값 + 파일 압축
        need_decompress = True
    
    if need_decompress:
        try:
            # zlib raw deflate (no header)
            image_data = zlib.decompress(image_data, -15)
        except zlib.error:
            # 압축되지 않은 경우 그대로 사용
            pass
    
    return image_data

def save_image(image_data: bytes, output_path: str) -> bool:
    """이미지 데이터를 파일로 저장"""
    try:
        with open(output_path, 'wb') as f:
            f.write(image_data)
        return True
    except:
        return False
```

### 3.6 전체 이미지 파싱 흐름

```python
def parse_picture_from_records(records: list, start_idx: int, doc_info: dict) -> dict:
    """
    레코드 리스트에서 그림 전체 파싱
    
    records: 전체 레코드 리스트
    start_idx: CTRL_HEADER 레코드 인덱스
    doc_info: DocInfo에서 파싱한 정보 (bin_data 포함)
    """
    base_level = records[start_idx]['level']
    common_attrs = None
    shape_component = None
    picture_attrs = None
    
    i = start_idx + 1
    while i < len(records):
        rec = records[i]
        tag_id = rec['tag_id']
        level = rec['level']
        data = rec['data']
        
        if level <= base_level:
            break
        
        if tag_id == HWPTAG_SHAPE_COMPONENT:
            shape_component = parse_shape_component(data)
            
        elif tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            picture_attrs = parse_picture(data)
        
        i += 1
    
    # bin_data_id로 DocInfo의 바이너리 정보 조회
    bin_data_id = picture_attrs.get('bin_data_id', 0) if picture_attrs else 0
    bin_data_info = None
    
    if bin_data_id > 0 and 'bin_data' in doc_info:
        # bin_data_id는 1부터 시작, 리스트는 0부터
        idx = bin_data_id - 1
        if 0 <= idx < len(doc_info['bin_data']):
            bin_data_info = doc_info['bin_data'][idx]
    
    return {
        'shape_component': shape_component,
        'picture_attrs': picture_attrs,
        'bin_data_id': bin_data_id,
        'bin_data_info': bin_data_info,
        'end_index': i,
    }
```

### 3.7 이미지 → 마크다운 변환

```python
def picture_to_markdown(picture_data: dict, 
                        ole: olefile.OleFileIO,
                        output_dir: str,
                        compressed: bool) -> str:
    """
    파싱된 그림 데이터를 마크다운 이미지로 변환
    이미지 파일 추출 후 마크다운 링크 생성
    """
    bin_data_id = picture_data.get('bin_data_id', 0)
    bin_data_info = picture_data.get('bin_data_info')
    
    if not bin_data_info or bin_data_id == 0:
        return '<!-- Image not found -->'
    
    # 이미지 추출
    image_data = extract_image_from_hwp(ole, bin_data_id, bin_data_info, compressed)
    
    if image_data is None:
        return '<!-- Image extraction failed -->'
    
    # 파일 저장
    extension = bin_data_info.get('extension', 'jpg').lower()
    filename = f'image_{bin_data_id:04d}.{extension}'
    filepath = os.path.join(output_dir, filename)
    
    if save_image(image_data, filepath):
        # alt 텍스트 생성
        shape = picture_data.get('shape_component', {})
        width_mm = shape.get('current_width_mm', 0)
        height_mm = shape.get('current_height_mm', 0)
        
        alt_text = f'Image {bin_data_id}'
        if width_mm and height_mm:
            alt_text = f'Image {bin_data_id} ({width_mm:.1f}mm x {height_mm:.1f}mm)'
        
        return f'![{alt_text}]({filename})'
    else:
        return '<!-- Image save failed -->'
```

---

## 4. 통합 파서 구현

### 4.1 전체 파싱 클래스

```python
import olefile
import zlib
import struct
import os

class HWPParser:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.ole = olefile.OleFileIO(filepath)
        self.header = None
        self.doc_info = None
        self.bin_data_list = []
        
    def parse(self) -> dict:
        """HWP 파일 전체 파싱"""
        # 1. FileHeader
        self.header = self._parse_file_header()
        
        if self.header['encrypted']:
            raise ValueError("암호화된 문서는 지원하지 않습니다")
        
        # 2. DocInfo
        self.doc_info = self._parse_doc_info()
        
        # 3. BodyText
        sections = self._parse_body_text()
        
        return {
            'header': self.header,
            'doc_info': self.doc_info,
            'sections': sections,
        }
    
    def _decompress(self, data: bytes) -> bytes:
        """압축 해제"""
        if self.header and self.header['compressed']:
            try:
                return zlib.decompress(data, -15)
            except:
                return data
        return data
    
    def _parse_file_header(self) -> dict:
        data = self.ole.openstream('FileHeader').read()
        signature = data[0:32]
        
        if not signature.startswith(b"HWP Document File"):
            raise ValueError("Invalid HWP file")
        
        version = struct.unpack('<I', data[32:36])[0]
        flags = struct.unpack('<I', data[36:40])[0]
        
        return {
            'version': version,
            'compressed': bool(flags & 0x01),
            'encrypted': bool(flags & 0x02),
            'distributed': bool(flags & 0x04),
        }
    
    def _parse_doc_info(self) -> dict:
        data = self.ole.openstream('DocInfo').read()
        data = self._decompress(data)
        records = self._parse_records(data)
        
        doc_info = {
            'bin_data': [],
            'char_shapes': [],
            'para_shapes': [],
            'styles': [],
        }
        
        for rec in records:
            tag_id = rec['tag_id']
            rec_data = rec['data']
            
            if tag_id == HWPTAG_BIN_DATA:
                doc_info['bin_data'].append(parse_bin_data_info(rec_data))
            elif tag_id == HWPTAG_CHAR_SHAPE:
                doc_info['char_shapes'].append(parse_char_shape(rec_data))
            elif tag_id == HWPTAG_PARA_SHAPE:
                doc_info['para_shapes'].append(parse_para_shape(rec_data))
        
        return doc_info
    
    def _parse_body_text(self) -> list:
        sections = []
        section_idx = 0
        
        while True:
            stream_name = f'BodyText/Section{section_idx}'
            if not self.ole.exists(stream_name):
                break
            
            data = self.ole.openstream(stream_name).read()
            data = self._decompress(data)
            records = self._parse_records(data)
            
            section = self._process_section_records(records)
            sections.append(section)
            section_idx += 1
        
        return sections
    
    def _parse_records(self, data: bytes) -> list:
        records = []
        pos = 0
        
        while pos < len(data):
            if pos + 4 > len(data):
                break
            
            header = struct.unpack('<I', data[pos:pos+4])[0]
            tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF
            pos += 4
            
            if size == 0xFFF:
                size = struct.unpack('<I', data[pos:pos+4])[0]
                pos += 4
            
            payload = data[pos:pos+size]
            pos += size
            
            records.append({
                'tag_id': tag_id,
                'level': level,
                'size': size,
                'data': payload,
            })
        
        return records
    
    def _process_section_records(self, records: list) -> dict:
        """섹션의 레코드들을 처리하여 구조화된 데이터 생성"""
        elements = []  # 문단, 표, 이미지 등
        
        i = 0
        while i < len(records):
            rec = records[i]
            tag_id = rec['tag_id']
            
            if tag_id == HWPTAG_PARA_HEADER:
                # 일반 문단 처리
                para = self._parse_paragraph(records, i)
                elements.append({'type': 'paragraph', 'data': para})
                i = para.get('end_index', i + 1)
                
            elif tag_id == HWPTAG_CTRL_HEADER:
                ctrl_id = struct.unpack('<I', rec['data'][0:4])[0]
                
                if ctrl_id == CTRL_ID_TABLE:
                    # 표 처리
                    table = parse_table_from_records(records, i)
                    elements.append({'type': 'table', 'data': table})
                    i = table.get('end_index', i + 1)
                    
                elif ctrl_id == CTRL_ID_PICTURE:
                    # 이미지 처리
                    picture = parse_picture_from_records(records, i, self.doc_info)
                    elements.append({'type': 'picture', 'data': picture})
                    i = picture.get('end_index', i + 1)
                else:
                    i += 1
            else:
                i += 1
        
        return {'elements': elements}
    
    def _parse_paragraph(self, records: list, start_idx: int) -> dict:
        """문단 파싱"""
        rec = records[start_idx]
        para_header = parse_para_header(rec['data'])
        base_level = rec['level']
        
        text = ''
        char_shapes = []
        
        i = start_idx + 1
        while i < len(records):
            rec = records[i]
            if rec['level'] <= base_level and rec['tag_id'] != HWPTAG_PARA_TEXT:
                break
            
            if rec['tag_id'] == HWPTAG_PARA_TEXT:
                text = parse_para_text(rec['data'])
            elif rec['tag_id'] == HWPTAG_PARA_CHAR_SHAPE:
                char_shapes = parse_para_char_shape(
                    rec['data'], 
                    para_header.get('char_shape_count', 0)
                )
            
            i += 1
        
        return {
            'header': para_header,
            'text': text,
            'char_shapes': char_shapes,
            'end_index': i,
        }
    
    def extract_image(self, bin_data_id: int) -> bytes:
        """이미지 바이너리 추출"""
        if bin_data_id <= 0 or bin_data_id > len(self.doc_info['bin_data']):
            return None
        
        bin_info = self.doc_info['bin_data'][bin_data_id - 1]
        return extract_image_from_hwp(
            self.ole, 
            bin_data_id, 
            bin_info, 
            self.header['compressed']
        )
    
    def close(self):
        self.ole.close()


# Tag ID 상수
HWPTAG_BEGIN = 0x010
HWPTAG_BIN_DATA = HWPTAG_BEGIN + 2
HWPTAG_CHAR_SHAPE = HWPTAG_BEGIN + 5
HWPTAG_PARA_SHAPE = HWPTAG_BEGIN + 9
HWPTAG_PARA_HEADER = HWPTAG_BEGIN + 50
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51
HWPTAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52
HWPTAG_CTRL_HEADER = HWPTAG_BEGIN + 55
HWPTAG_LIST_HEADER = HWPTAG_BEGIN + 56
HWPTAG_SHAPE_COMPONENT = HWPTAG_BEGIN + 60
HWPTAG_TABLE = HWPTAG_BEGIN + 61
HWPTAG_SHAPE_COMPONENT_PICTURE = HWPTAG_BEGIN + 69

# Control ID 상수
CTRL_ID_TABLE = 0x74626C20    # 'tbl '
CTRL_ID_PICTURE = 0x24706963  # '$pic'
```

---

## 5. 주요 주의사항

### 5.1 레코드 Level 처리

```
★ 핵심: Level이 감소하면 현재 개체 종료

예: 표 다음에 문단이 오는 경우
[CTRL_HEADER level=0]  ← 표 시작
  [TABLE level=1]
  [LIST_HEADER level=2]
    [PARA_HEADER level=3]
    [PARA_TEXT level=4]
[PARA_HEADER level=0]  ← level 0으로 돌아옴 = 표 종료, 새 문단 시작
```

### 5.2 BinData ID 주의

```python
# DocInfo의 bin_data 리스트: 0-indexed
# 그림 개체의 bin_data_id: 1-indexed

# 변환: bin_data_id → 리스트 인덱스
list_index = bin_data_id - 1

# BinData 스트림 이름: 16진수 4자리
stream_name = f'BinData/BIN{bin_data_id:04X}.{extension}'
# 예: bin_data_id=1 → BIN0001.jpg
```

### 5.3 셀 병합 처리

```python
# 셀 속성의 col_span, row_span은 "병합된 셀 수"
# col_span=2면 현재 셀 + 오른쪽 1개 = 총 2칸 차지
# row_span=3이면 현재 셀 + 아래 2개 = 총 3칸 차지

# 마크다운은 병합 미지원 → 첫 셀에만 내용, 나머지 빈칸
# HTML 변환 시 colspan, rowspan 속성 사용
```

### 5.4 그림 개체 레코드 순서

```
CTRL_HEADER (ctrl_id = '$pic')
  └── SHAPE_COMPONENT (개체 요소 속성)
        └── SHAPE_COMPONENT_PICTURE (그림 속성 + bin_data_id)
```

### 5.5 압축 해제

```python
# zlib.decompress(data, -15)
# -15 = raw deflate (no zlib/gzip header)
# 표준 zlib 헤더가 없으므로 wbits에 음수 사용
```
