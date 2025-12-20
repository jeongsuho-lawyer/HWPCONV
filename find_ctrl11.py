"""PARA_TEXT에서 제어문자 11 찾기"""
import struct
import olefile
import zlib

file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

pos = 0
para_text_idx = 0
ctrl_11_count = 0

print("=== PARA_TEXT 안의 제어문자 11 찾기 ===\n")

while pos < len(section_data):
    if pos + 4 > len(section_data):
        break
    
    header = struct.unpack('<I', section_data[pos:pos+4])[0]
    tag_id = header & 0x3FF
    size = (header >> 20) & 0xFFF
    pos += 4
    
    if size == 0xFFF:
        size = struct.unpack('<I', section_data[pos:pos+4])[0]
        pos += 4
    
    # Normalize
    if 66 <= tag_id <= 77:
        if tag_id == 72:
            normalized = 24
        else:
            normalized = tag_id - 50
    else:
        normalized = tag_id
    
    if normalized == 17:  # PARA_TEXT
        payload = section_data[pos:pos+size]
        
        # UTF-16LE로 문자 읽기
        text_pos = 0
        found_ctrl_11 = False
        
        while text_pos < len(payload):
            if text_pos + 2 > len(payload):
                break
            
            char_code = struct.unpack('<H', payload[text_pos:text_pos+2])[0]
            
            if char_code == 11:  # 제어문자 11!
                found_ctrl_11 = True
                ctrl_11_count += 1
                print(f"PARA_TEXT #{para_text_idx}: 제어문자 11 발견!")
                print(f"  위치: {text_pos} / {len(payload)} bytes")
                
                # 다음 16바이트 (8 WCHAR) 확인
                if text_pos + 16 <= len(payload):
                    ctrl_data = payload[text_pos:text_pos+16]
                    print(f"  제어 데이터: {ctrl_data.hex()}")
                print()
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
        
        para_text_idx += 1
    
    pos += size

print(f"\n총 PARA_TEXT: {para_text_idx}개")
print(f"제어문자 11 발견: {ctrl_11_count}개")

ole.close()
