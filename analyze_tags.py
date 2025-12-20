"""모든 태그 타입 분석"""
from hwpconv.parsers.hwp import HwpParser
import struct
import olefile
import zlib
from collections import Counter

file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

tag_counts = Counter()
pos = 0

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
    
    tag_counts[normalized] += 1
    pos += size

print("=== 모든 태그 분포 ===")
for tag_id, count in sorted(tag_counts.items()):
    tag_name = {
        16: 'PARA_HEADER',
        17: 'PARA_TEXT', 
        18: 'PARA_CHAR_SHAPE',
        19: 'PARA_LINE_SEG',
        21: 'CTRL_HEADER',
        24: 'LIST_HEADER',
        77: 'TABLE',
        5: 'CHAR_SHAPE',
        9: 'PARA_SHAPE',
    }.get(tag_id, f'TAG_{tag_id}')
    print(f"{tag_name:20s} (ID {tag_id:3d}): {count:4d}개")

ole.close()
