"""HWP 파서 완전 재작성 - 사양서 기반

주요 수정사항:
1. _extract_text_with_ctrls: 제어문자 완전 구현
2. _parse_ctrl_header: CTRL_HEADER 파싱 (이미지/표 위치)
3. _parse_table: TABLE 완전 파싱
4. _parse_list_header: LIST_HEADER 파싱 (셀)
5. _parse_para_shape: 정렬/들여쓰기
6. 이미지 위치 매핑

구현 순서:
1. 제어문자 처리 로직
2. CTRL_HEADER → 이미지/표 식별
3. TABLE + LIST_HEADER → 표 파싱
4. PARA_SHAPE → 정렬 적용
"""

# 1단계: _extract_text_with_ctrls 완전 재작성
def _extract_text_with_ctrls_NEW(data, char_count=0):
    \"\"\"사양서 기준 완전 구현\"\"\"
    chars = []
    ctrl_info = []
    pos = 0
    char_pos = 0
    
    CHAR_CTRLS = {0, 10, 13, 24, 30, 31}
    INLINE_CTRLS = {4, 9}
    EXTENDED_CTRLS = {2, 3, 11, 15, 16, 17, 18, 21, 22, 23}
    
    while pos < len(data):
        if char_count > 0 and char_pos >= char_count:
            break
        if pos + 2 > len(data):
            break
        
        code = struct.unpack('<H', data[pos:pos+2])[0]
        
        if code < 32:
            if code in CHAR_CTRLS:
                # 1 WCHAR
                if code == 10:
                    chars.append('\\n')
                elif code == 30:
                    chars.append('\\u00A0')
                elif code == 31:
                    chars.append(' ')
                ctrl_info.append((char_pos, code, b''))
                pos += 2
                char_pos += 1
            
            elif code in INLINE_CTRLS:
                # 8 WCHAR, 위치 1개
                ctrl_data = data[pos:pos+16]
                if code == 9:
                    chars.append('\\t')
                ctrl_info.append((char_pos, code, ctrl_data))
                pos += 16
                char_pos += 1
            
            elif code in EXTENDED_CTRLS:
                # 8 WCHAR, 위치 차지 안함
                ctrl_data = data[pos:pos+16]
                ctrl_info.append((char_pos, code, ctrl_data))
                pos += 16
                # char_pos 증가 안함!
            
            else:
                pos += 2
                char_pos += 1
        else:
            chars.append(chr(code))
            pos += 2
            char_pos += 1
    
    return ''.join(chars), ctrl_info


# 2단계: _parse_ctrl_header 새로 추가
def _parse_ctrl_header_NEW(data):
    \"\"\"CTRL_HEADER 파싱\"\"\"
    if len(data) < 4:
        return None
    
    ctrl_id = struct.unpack('<I', data[0:4])[0]
    ctrl_ch = ''.join([
        chr((ctrl_id >> 24) & 0xFF),
        chr((ctrl_id >> 16) & 0xFF),
        chr((ctrl_id >> 8) & 0xFF),
        chr(ctrl_id & 0xFF)
    ])
    
    return {
        'ctrl_id': ctrl_id,
        'ctrl_ch': ctrl_ch,
        'data': data[4:]
    }


# 3단계: _parse_table 완전 재작성
def _parse_table_NEW(data):
    \"\"\"TABLE 레코드 파싱 (사양서 기준)\"\"\"
    if len(data) < 18:
        return None
    
    attr = struct.unpack('<I', data[0:4])[0]
    row_count = struct.unpack('<H', data[4:6])[0]
    col_count = struct.unpack('<H', data[6:8])[0]
    cell_spacing = struct.unpack('<H', data[8:10])[0]
    margins = struct.unpack('<4H', data[10:18])
    
    # 행 높이
    offset = 18
    row_heights = []
    for _ in range(row_count):
        if offset + 2 > len(data):
            break
        h = struct.unpack('<H', data[offset:offset+2])[0]
        row_heights.append(h)
        offset += 2
    
    return {
        'rows': row_count,
        'cols': col_count,
        'cell_spacing': cell_spacing,
        'margins': margins,
        'row_heights': row_heights
    }


# 4단계: _parse_list_header
def _parse_list_header_NEW(data):
    \"\"\"LIST_HEADER 파싱 (셀)\"\"\"
    if len(data) < 6:
        return None
    
    para_count = struct.unpack('<h', data[0:2])[0]
    attr = struct.unpack('<I', data[2:6])[0]
    
    text_direction = attr & 0x07
    vert_align = (attr >> 5) & 0x03
    
    return {
        'para_count': para_count,
        'text_direction': text_direction,
        'vert_align': vert_align
    }
