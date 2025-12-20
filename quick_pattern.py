"""제어문자 11 → LIST_HEADER 패턴 확인 (간결)"""
from hwpconv.parsers.hwp import HwpParser
import struct

parser = HwpParser()
file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

import olefile, zlib

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)
records = parser._parse_records(section_data)

print("제어문자 11 → 다음 레코드 패턴:\n")

count = 0
for idx, (tag_id, data, level) in enumerate(records):
    if tag_id == 17:  # PARA_TEXT
        # 제어문자 11 찾기
        for text_pos in range(0, len(data), 2):
            if text_pos + 2 <= len(data):
                char_code = struct.unpack('<H', data[text_pos:text_pos+2])[0]
                if char_code == 11:
                    # 다음 5개 레코드
                    next_tags = []
                    for j in range(idx+1, min(idx+6, len(records))):
                        next_tag = records[j][0]
                        next_level = records[j][2]
                        tag_name = {16:'PARA_H', 17:'PARA_T', 18:'CHAR_S', 21:'CTRL', 24:'LIST'}.get(next_tag, str(next_tag))
                        next_tags.append(f"{tag_name}(L{next_level})")
                    
                    print(f"[{idx}] {' → '.join(next_tags[:5])}")
                    
                    count += 1
                    if count >= 5:  # 처음 5개만
                        break
                    break
    if count >= 5:
        break

ole.close()
