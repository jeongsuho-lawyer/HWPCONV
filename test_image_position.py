"""190820 HWP 이미지 위치 확인"""
import sys
sys.path.insert(0, 'src')
import olefile, zlib, struct

ole = olefile.OleFileIO('190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp')

# 1. BinData 이미지 확인
print("=== BinData 이미지 ===")
images = {}
for entry in ole.listdir():
    if entry[0] == 'BinData':
        fname = entry[1]
        print(f"  {fname}")
        # BIN0001 -> image_id
        if fname.startswith('BIN'):
            hex_id = fname[3:7]
            bin_id = int(hex_id, 16)
            images[f'BIN{bin_id:04X}'] = fname

print(f"\n총 {len(images)}개 이미지: {list(images.keys())}")

# 2. Section에서 $pic CTRL_HEADER 찾기
print("\n=== $pic CTRL_HEADER 위치 ===")
section_data = ole.openstream('BodyText/Section0').read()
section_data = zlib.decompress(section_data, -15)

# 레코드 파싱
records = []
pos = 0
while pos < len(section_data) - 4:
    header = struct.unpack('<I', section_data[pos:pos+4])[0]
    pos += 4
    tag_id = header & 0x3FF
    level = (header >> 10) & 0x3FF
    size = (header >> 20) & 0xFFF
    if size == 0xFFF:
        size = struct.unpack('<I', section_data[pos:pos+4])[0]
        pos += 4
    record_data = section_data[pos:pos+size]
    pos += size
    
    # Normalize
    if 66 <= tag_id <= 77:
        tag_id = 24 if tag_id == 72 else tag_id - 50
    
    records.append((tag_id, record_data, level))

# $pic 찾기
pic_count = 0
for i, (tag_id, data, level) in enumerate(records):
    if tag_id == 21 and len(data) >= 4:  # CTRL_HEADER
        ctrl_id = struct.unpack('<I', data[0:4])[0]
        if ctrl_id == 0x24706963:  # '$pic'
            pic_count += 1
            print(f"  Record {i}: $pic (level={level})")

print(f"\n총 {pic_count}개 $pic 컨트롤")

ole.close()
