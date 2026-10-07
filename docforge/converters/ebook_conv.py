"""电子书格式转换：EPUB → TXT / MD / HTML / PDF / DOCX。

底层使用 ebooklib 读取 EPUB 并提取正文文本。
"""

from __future__ import annotations

import html as html_lib
import re
from pathlib import Path

from ebooklib import ITEM_DOCUMENT, epub

from docforge.converters.base import ConversionError, register
from docforge.converters.html_pdf import html_to_pdf, text_to_pdf
from docforge.converters.text_conv import _html_to_docx, _text_to_md


def _read_epub(src: Path) -> tuple[str, str]:
    """返回 (标题, 正文 HTML)。"""
    try:
        book = epub.read_epub(str(src))
    except Exception as e:
        raise ConversionError(f"无法读取 EPUB：{e}")

    title = ""
    if book.get_metadata("DC", "title"):
        title = book.get_metadata("DC", "title")[0][0] or ""

    parts = []
    for item in book.get_items_of_type(ITEM_DOCUMENT):
        try:
            content = item.get_content().decode("utf-8", errors="replace")
            parts.append(content)
        except Exception:
            continue
    if not parts:
        raise ConversionError("EPUB 中未找到正文内容")
    return title, "\n".join(parts)


def _epub_to_text(src: Path) -> str:
    _, html_text = _read_epub(src)
    text = re.sub(r"<script.*?</script>", "", html_text, flags=re.S | re.I)
    text = re.sub(r"<style.*?</style>", "", text, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", "\n", text)
    text = html_lib.unescape(text)
    lines = [l.strip() for l in text.splitlines()]
    return "\n".join(l for l in lines if l)


@register("epub", "txt")
def epub_to_txt(src: Path, dst: Path, options=None):
    dst.write_text(_epub_to_text(src), encoding="utf-8")


@register("epub", "md")
def epub_to_md(src: Path, dst: Path, options=None):
    dst.write_text(_text_to_md(_epub_to_text(src)), encoding="utf-8")


@register("epub", "html")
def epub_to_html(src: Path, dst: Path, options=None):
    title, html_text = _read_epub(src)
    full = (
        "<html><head><meta charset='utf-8'>"
        f"<title>{html_lib.escape(title or src.stem)}</title></head>"
        f"<body style='max-width:820px;margin:0 auto;padding:24px;"
        f"font-family:SimSun,serif;line-height:1.8'>{html_text}</body></html>"
    )
    dst.write_text(full, encoding="utf-8")


@register("epub", "pdf")
def epub_to_pdf(src: Path, dst: Path, options=None):
    title, html_text = _read_epub(src)
    html_to_pdf(html_text, dst, title=title or src.stem)


@register("epub", "docx")
def epub_to_docx(src: Path, dst: Path, options=None):
    _, html_text = _read_epub(src)
    _html_to_docx(html_text, dst)
