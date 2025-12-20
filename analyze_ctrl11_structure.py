"""제어문자 11 이후 레코드 구조 분석"""
from hwpconv.parsers.hwp import HwpParser
import struct

parser = HwpParser()
file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

import olefile
import zlib

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

records = parser._parse_records(section_data)

print(f"=== 제어문자 11 이후 레코드 구조 분석 ===\n")

for idx, (tag_id, data, level) in enumerate(records):
    if tag_id == parser.HWPTAG_PARA_TEXT:
        # PARA_TEXT에서 제어문자 11 찾기
        text_pos = 0
        has_ctrl_11 = False
        
        while text_pos < len(data):
            if text_pos + 2 > len(data):
                break
            
            char_code = struct.unpack('<H', data[text_pos:text_pos+2])[0]
            
            if char_code == 11:
                has_ctrl_11 = True
                print(f"\n[{idx}] PARA_TEXT에 제어문자 11 발견!")
                print(f"  level={level}")
                
                # 다음 10개 레코드 확인
                print(f"\n  다음 레코드들:")
                for j in range(idx+1, min(idx+11, len(records))):
                    next_tag, next_data, next_level = records[j]
                    
                    tag_name = {
                        16: 'PARA_HEADER',
                        17: 'PARA_TEXT',
                        18: 'PARA_CHAR_SHAPE',
                        19: 'PARA_LINE_SEG',
                        21: 'CTRL_HEADER',
                        24: 'LIST_HEADER',
                        77: 'TABLE',
                    }.get(next_tag, f'TAG_{next_tag}')
                    
                    print(f"    [{j}] {tag_name:20s} level={next_level:2d} size={len(next_data):5d}")
                    
                    # CTRL_HEADER면 ctrl_id 확인
                    if next_tag == 21 and len(next_data) >= 4:
                        ctrl_id = struct.unpack('<I', next_data[0:4])[0]
                        ctrl_ch = ''.join([
                            chr((ctrl_id >> 24) & 0xFF),
                            chr((ctrl_id >> 16) & 0xFF),
                            chr((ctrl_id >> 8) & 0xFF),
                            chr(ctrl_id & 0xFF)
                        ])
                        print(f"         ctrl_id=0x{ctrl_id:08X} '{ctrl_ch}'")
                    
                    # LIST_HEADER면 셀 정보 확인
                    if next_tag == 24:
                        print(f"         ★ LIST_HEADER (표 셀?))")
                
                break
            
            # 제어문자 처리
            if char_code < 32:
                if char_code in {0, 10, 13, 24, 30, 31}:
                    text_pos += 2
                elif char_code in {4, 9}:
                    text_pos += 16
                else:  # extended
                    text_pos += 16
            else:
                text_pos += 2
        
        if has_ctrl_11:
            # 최대 3개 패턴만 출력
            pass

ole.close()
