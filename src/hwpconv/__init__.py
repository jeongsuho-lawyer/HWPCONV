"""
hwpconv - HWP/HWPX → Markdown/HTML 변환기

한글 문서(.hwp, .hwpx)를 Markdown/HTML/Text로 변환하는 Python 라이브러리 및 CLI

Example:
    >>> from hwpconv import HwpxParser, HwpParser, MarkdownConverter
    >>> doc = HwpxParser().parse('document.hwpx')
    >>> md = MarkdownConverter().convert(doc)
    >>> print(md)
"""

__version__ = "0.1.0"

from .models import (
    Document,
    Section,
    Paragraph,
    TextRun,
    TextStyle,
    Table,
    TableRow,
    TableCell,
    Footnote,
    HeadingLevel,
    Image,
)
from .parsers.hwpx import HwpxParser
from .parsers.hwp import HwpParser
from .converters.markdown import MarkdownConverter
from .converters.html import HtmlConverter

__all__ = [
    # Version
    "__version__",
    # Models
    "Document",
    "Section",
    "Paragraph",
    "TextRun",
    "TextStyle",
    "Table",
    "TableRow",
    "TableCell",
    "Footnote",
    "HeadingLevel",
    "Image",
    # Parsers
    "HwpxParser",
    "HwpParser",
    # Converters
    "MarkdownConverter",
    "HtmlConverter",
]
