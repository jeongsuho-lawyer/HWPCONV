"""MD 변환 테스트"""
import sys
sys.path.insert(0, 'src')
from hwpconv.parsers.hwp import HwpParser
from hwpconv.converters.markdown import MarkdownConverter

class NoAPIParser(HwpParser):
    def _extract_images(self, ole, doc):
        pass

p = NoAPIParser()
hwp_file = '190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp'
doc = p.parse(hwp_file)

md = MarkdownConverter().convert(doc)
with open('test_190820_fixed.md', 'w', encoding='utf-8') as f:
    f.write(md)

print(f"Lines: {len(md.splitlines())}")
print(f"Chars: {len(md)}")

# 누락 테스트
if "붙임1" in md:
    print("✓ 붙임1 포함")
else:
    print("✗ 붙임1 없음")

if "시간 계획" in md:
    print("✓ 시간 계획 포함")
else:
    print("✗ 시간 계획 없음")
