"""CTRL_HEADER 분석 - 표 찾기"""
import struct
import olefile
import zlib

file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

pos = 0
ctrl_idx = 0

print("=== CTRL_HEADER 분석 ===\n")

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
    
    if normalized == 21:  # CTRL_HEADER
        payload = section_data[pos:pos+size]
        if len(payload) >= 4:
            ctrl_id = struct.unpack('<I', payload[0:4])[0]
            
            # ctrl_id를 4글자로 변환
            ctrl_ch = ''.join([
                chr((ctrl_id >> 24) & 0xFF),
                chr((ctrl_id >> 16) & 0xFF),
                chr((ctrl_id >> 8) & 0xFF),
                chr(ctrl_id & 0xFF)
            ])
            
            print(f"CTRL_HEADER #{ctrl_idx}:")
            print(f"  ctrl_id: 0x{ctrl_id:08X}")
            print(f"  ctrl_ch: '{ctrl_ch}'")
            print(f"  size: {size} bytes")
            
            if ctrl_ch == 'tbl ' or ctrl_ch.strip() == 'tbl':
                print(f"  ★ 표 발견! ★")
            print()
            
            ctrl_idx += 1
    
    pos += size

print(f"\n총 CTRL_HEADER: {ctrl_idx}개")

ole.close()
