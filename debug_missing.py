"""표 파싱 디버깅 - 왜 Lines 107인지"""
import sys
sys.path.insert(0, 'src')
from hwpconv.parsers.hwp import HwpParser
import olefile, zlib, struct

ole = olefile.OleFileIO('190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp')
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

p = HwpParser()
records = p._parse_records(section_data)

# 표 파싱 시뮬레이션
print("=== 표 파싱 시뮬레이션 ===")

# 첫 번째 표 (idx 12)
start_idx = 12
base_level = records[start_idx][2]
print(f"표 시작: idx={start_idx}, base_level={base_level}")

table, consumed = p._parse_table_from_records_v2(records, start_idx)
print(f"반환: consumed={consumed}, {table.row_count}x{table.col_count}")

# 두 번째 표 확인
next_idx = start_idx + consumed
print(f"다음 인덱스: {next_idx}")
if next_idx < len(records):
    tag_id, data, level = records[next_idx]
    extra = ""
    if tag_id == p.HWPTAG_CTRL_HEADER and len(data) >= 4:
        ctrl_id = struct.unpack('<I', data[0:4])[0]
        extra = f" ctrl=0x{ctrl_id:08X}"
    print(f"다음 레코드: tag={tag_id}, level={level}{extra}")

ole.close()
