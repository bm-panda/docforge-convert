"""PDF 相关转换：PDF → TXT / MD / HTML / DOCX / PNG / JPG。

底层使用 PyMuPDF（pymupdf）提取文本与渲染页面。
PDF→DOCX / PDF→HTML 优先使用 pdf2docx（保留布局、表格与图片），
pdf2docx 不可用时回退到文本提取方案。
扫描件（无文字层）PDF：若系统装有 Tesseract OCR 则自动 OCR，否则给出明确提示。
"""

from __future__ import annotations

import contextlib
import html as html_lib
import base64
import io
import logging
from pathlib import Path

import pymupdf

from docforge.converters.base import ConversionError, register
from docforge.converters.text_conv import _text_to_docx, _text_to_md
from docforge.core.utils import unique_output_path


@contextlib.contextmanager
def _silence_pdf2docx():
    """静音 pdf2docx 的 INFO 日志输出（其通过根 logger 打印）。"""
    root = logging.getLogger()
    old_level = root.level
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        root.setLevel(logging.CRITICAL)
        try:
            yield
        finally:
            root.setLevel(old_level)


def _open_pdf(src: Path) -> pymupdf.Document:
    try:
        return pymupdf.open(str(src))
    except Exception as e:
        raise ConversionError(f"无法打开 PDF：{e}")


def _has_text_layer(doc: pymupdf.Document) -> bool:
    for page in doc:
        if page.get_text("text").strip():
            return True
    return False


def _try_ocr(doc: pymupdf.Document, lang: str = "chi_sim+eng") -> str:
    """扫描件 OCR（需要系统安装 Tesseract）。失败返回空字符串。"""
    try:
        import pytesseract
        from PIL import Image
        import shutil
        exe = shutil.which("tesseract")
        if not exe:
            return ""
        pytesseract.pytesseract.tesseract_cmd = exe
        parts = []
        for page in doc:
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2.5, 2.5))
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            parts.append(pytesseract.image_to_string(img, lang=lang))
        return "\n".join(parts)
    except Exception:
        return ""


def _pdf_extract_text(src: Path) -> str:
    doc = _open_pdf(src)
    try:
        text = "\n".join(page.get_text("text") for page in doc)
        if text.strip():
            return text
        # 无文字层 → 尝试 OCR
        ocr = _try_ocr(doc)
        if ocr.strip():
            return ocr
        raise ConversionError(
            "PDF 中未提取到文本（该 PDF 可能为扫描件/图片型）。\n"
            "可改用 PDF→PNG 导出为图片；或在系统安装 Tesseract OCR 后自动识别文字。"
        )
    finally:
        doc.close()


def _pdf_has_scan_no_text(src: Path) -> bool:
    doc = _open_pdf(src)
    try:
        return not _has_text_layer(doc)
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# PDF → DOCX（保布局/图片/表格）
# ---------------------------------------------------------------------------

@register("pdf", "docx")
def pdf_to_docx(src: Path, dst: Path, options=None):
    try:
        from pdf2docx import Converter
        with _silence_pdf2docx():
            cv = Converter(str(src))
            try:
                cv.convert(str(dst))
            finally:
                cv.close()
        if Path(dst).exists() and Path(dst).stat().st_size > 0:
            return
    except Exception:
        pass
    # 回退：纯文本
    text = _pdf_extract_text(src)
    _text_to_docx(text, dst)


# ---------------------------------------------------------------------------
# PDF → HTML
# ---------------------------------------------------------------------------

@register("pdf", "html")
def pdf_to_html(src: Path, dst: Path, options=None):
    try:
        from pdf2docx import Converter
        with _silence_pdf2docx():
            cv = Converter(str(src))
            try:
                cv.convert_html(str(dst))
            finally:
                cv.close()
        if Path(dst).exists() and Path(dst).stat().st_size > 0:
            return
    except Exception:
        pass
    # 回退：逐页文本 + 内嵌页面图片
    doc = _open_pdf(src)
    try:
        parts = ["<html><head><meta charset='utf-8'>",
                 f"<title>{html_lib.escape(src.stem)}</title></head>",
                 "<body style='font-family:SimSun,serif;line-height:1.7'>"]
        for i, page in enumerate(doc):
            text = page.get_text("text")
            if text.strip():
                for line in text.splitlines():
                    if line.strip():
                        parts.append(f"<p>{html_lib.escape(line)}</p>")
            # 内嵌页面图片（扫描件场景）
            pix = page.get_pixmap(matrix=pymupdf.Matrix(1.2, 1.2))
            data = pix.tobytes("png")
            b64 = base64.b64encode(data).decode("ascii")
            parts.append(f'<p><img src="data:image/png;base64,{b64}" '
                         f"style='max-width:100%'/></p>")
            parts.append("<hr/>")
        parts.append("</body></html>")
        dst.write_text("".join(parts), encoding="utf-8")
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# PDF → TXT / MD
# ---------------------------------------------------------------------------

@register("pdf", "txt")
def pdf_to_txt(src: Path, dst: Path, options=None):
    dst.write_text(_pdf_extract_text(src), encoding="utf-8")


@register("pdf", "md")
def pdf_to_md(src: Path, dst: Path, options=None):
    dst.write_text(_text_to_md(_pdf_extract_text(src)), encoding="utf-8")


# ---------------------------------------------------------------------------
# PDF → 图片
# ---------------------------------------------------------------------------

def _render_pdf_to_images(src: Path, dst: Path, fmt: str, dpi: int = 150):
    doc = _open_pdf(src)
    try:
        if doc.page_count == 0:
            raise ConversionError("PDF 为空文档")
        zoom = dpi / 72.0
        mat = pymupdf.Matrix(zoom, zoom)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=mat)
            if i == 0:
                out = dst
            else:
                out = unique_output_path(dst.parent, f"{dst.stem}_p{i+1}", fmt)
            pix.pil_save(str(out), format=fmt.upper())
    finally:
        doc.close()


@register("pdf", "png")
def pdf_to_png(src: Path, dst: Path, options=None):
    _render_pdf_to_images(src, dst, "png")


@register("pdf", "jpg")
def pdf_to_jpg(src: Path, dst: Path, options=None):
    _render_pdf_to_images(src, dst, "jpeg")
