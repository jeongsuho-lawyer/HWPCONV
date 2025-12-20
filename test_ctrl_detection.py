"""간단 테스트: CTRL_HEADER=21 매칭 확인"""
from hwpconv.parsers.hwp import HwpParser
import olefile
import zlib
import struct

parser = HwpParser()
file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

# _parse_records 사용
records = parser._parse_records(section_data)

print(f"Total records: {len(records)}\n")

ctrl_header_count = 0
table_ctrl_count = 0

for idx, (tag_id, data, level) in enumerate(records):
    if tag_id == parser.HWPTAG_CTRL_HEADER:
        ctrl_header_count += 1
        if len(data) >= 4:
            ctrl_id = struct.unpack('<I', data[0:4])[0]
            print(f"[{idx}] CTRL_HEADER: ctrl_id=0x{ctrl_id:08X}, level={level}")
            
            if ctrl_id == 0x74626C20:
                print(f"     ★ TABLE CTRL_ID 발견!")
                table_ctrl_count += 1
                
                # 다음 몇 개 레코드 확인
                print(f"     다음 레코드들:")
                for j in range(idx+1, min(idx+5, len(records))):
                    next_tag, next_data, next_level = records[j]
                    print(f"       [{j}] tag={next_tag}, level={next_level}, size={len(next_data)}")

print(f"\n총 CTRL_HEADER: {ctrl_header_count}")
print(f"TABLE ctrl_id 발견: {table_ctrl_count}")
print(f"HWPTAG_CTRL_HEADER 상수값: {parser.HWPTAG_CTRL_HEADER}")

ole.close()
