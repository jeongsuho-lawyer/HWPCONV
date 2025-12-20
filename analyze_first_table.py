"""첫 번째 표 구조 분석"""
from hwpconv.parsers.hwp import HwpParser
import olefile, zlib

parser = HwpParser()
file_path = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'

ole = olefile.OleFileIO(file_path)
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)
records = parser._parse_records(section_data)

print("=== 첫 번째 표 구조 분석 ===\n")

# 첫 번째 ctrl 11 찾기
first_table_idx = None
for idx, (tag_id, data, level) in enumerate(records):
    if tag_id == 16:  # PARA_HEADER
        para_info = parser._parse_para_header(data)
        
        if idx + 1 < len(records) and records[idx + 1][0] == 17:
            text, ctrl_info = parser._extract_text_with_ctrls(records[idx + 1][1], para_info.get('char_count', 0))
            
            for pos, ctrl_code, ctrl_data in ctrl_info:
                if ctrl_code == 11:
                    first_table_idx = idx
                    print(f"첫 번째 ctrl 11 at PARA idx={idx}, level={level}")
                    
                    # 다음 20개 레코드 확인
                    print(f"\n다음 레코드들:")
                    list_header_count = 0
                    for j in range(idx + 1, min(idx + 30, len(records))):
                        next_tag, next_data, next_level = records[j]
                        tag_name = {16:'PARA_H', 17:'PARA_T', 18:'CHAR_S', 19:'LINE_S', 24:'LIST', 21:'CTRL'}.get(next_tag, str(next_tag))
                        
                        if next_tag == 24:
                            list_header_count += 1
                            # LIST_HEADER 내부 첫 문단 텍스트 확인
                            if j + 1 < len(records) and records[j + 1][0] == 16:
                                if j + 2 < len(records) and records[j + 2][0] == 17:
                                    ph = parser._parse_para_header(records[j + 1][1])
                                    t, _ = parser._extract_text_with_ctrls(records[j + 2][1], ph.get('char_count', 0))
                                    preview = t[:50].replace('\n', ' ') if t else ''
                                    print(f"  [{j}] {tag_name:8s} level={next_level} → '{preview}...'")
                                else:
                                    print(f"  [{j}] {tag_name:8s} level={next_level}")
                            else:
                                print(f"  [{j}] {tag_name:8s} level={next_level}")
                        else:
                            print(f"  [{j}] {tag_name:8s} level={next_level}")
                    
                    print(f"\n총 LIST_HEADER: {list_header_count}개")
                    break
            
            if first_table_idx:
                break
    
    if first_table_idx:
        break

ole.close()
