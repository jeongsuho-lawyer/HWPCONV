from pypdf import PdfReader
from PIL import Image
import io

# PDF를 이미지로 변환
pdf_path = '한글문서파일형식_5.0_revision1.3.pdf'
reader = PdfReader(pdf_path)

# 페이지 24 (0-indexed로 23)
page = reader.pages[23]

# PDF 페이지를 이미지로 렌더링
# PyPDF는 직접 렌더링을 지원하지 않으므로 다른 방법 필요
print("PyPDF로는 이미지 변환 불가")
print("대신 수동으로 PDF 페이지 24의 스크린샷을 요청해야 합니다")
