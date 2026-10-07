"""旧版 Office 格式转换：DOC、PPT（需要安装 MS Office）。

底层通过 COM 调用 Word / PowerPoint 完成高保真转换；
未安装 Office 时抛出 needs_office 错误，由上层提示。
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from docforge.converters.base import ConversionError, register
from docforge.converters import office_com
from docforge.converters.office_conv import _pptx_extract
from docforge.converters.text_conv import _text_to_md
from docforge.core.utils import unique_output_path


def _require_office() -> None:
    if not office_com._have_com():
        raise ConversionError("该格式需要安装 Microsoft Office 才能转换", needs_office=True)


def _ensure(ok: bool, msg: str):
    if not ok:
        raise ConversionError(msg, needs_office=True)


# ---------------------------------------------------------------------------
# DOC → 各类
# ---------------------------------------------------------------------------

@register("doc", "pdf")
def doc_to_pdf(src: Path, dst: Path, options=None):
    _require_office()
    _ensure(office_com.word_to_pdf(src, dst), "Word 转换 PDF 失败")


@register("doc", "docx")
def doc_to_docx(src: Path, dst: Path, options=None):
    _require_office()
    _ensure(office_com.word_save_as(src, dst, 12), "Word 转换 DOCX 失败")  # wdFormatXMLDocument


@register("doc", "txt")
def doc_to_txt(src: Path, dst: Path, options=None):
    _require_office()
    _ensure(office_com.word_save_as(src, dst, 7), "Word 转换 TXT 失败")  # wdFormatUnicodeText


@register("doc", "md")
def doc_to_md(src: Path, dst: Path, options=None):
    _require_office()
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "out.txt"
        _ensure(office_com.word_save_as(src, tmp, 7), "Word 转换失败")
        text = tmp.read_text(encoding="utf-8", errors="replace")
        dst.write_text(_text_to_md(text), encoding="utf-8")


@register("doc", "html")
def doc_to_html(src: Path, dst: Path, options=None):
    _require_office()
    _ensure(office_com.word_save_as(src, dst, 8), "Word 转换 HTML 失败")  # wdFormatHTML


@register("doc", "rtf")
def doc_to_rtf(src: Path, dst: Path, options=None):
    _require_office()
    _ensure(office_com.word_save_as(src, dst, 6), "Word 转换 RTF 失败")  # wdFormatRTF


# ---------------------------------------------------------------------------
# PPT → 各类
# ---------------------------------------------------------------------------

@register("ppt", "pdf")
def ppt_to_pdf(src: Path, dst: Path, options=None):
    _require_office()
    _ensure(office_com.ppt_to_pdf(src, dst), "PowerPoint 转换 PDF 失败")


@register("ppt", "pptx")
def ppt_to_pptx(src: Path, dst: Path, options=None):
    _require_office()
    import win32com.client
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            pres = app.Presentations.Open(str(src.resolve()), ReadOnly=True, WithWindow=False)
            try:
                pres.SaveAs(str(dst.resolve()), 24)  # ppSaveAsOpenXMLPresentation
                if not (Path(dst).exists() and Path(dst).stat().st_size > 0):
                    raise ConversionError("PowerPoint 转换 PPTX 失败", needs_office=True)
            finally:
                pres.Close()
        finally:
            app.Quit()
    except ConversionError:
        raise
    except Exception as e:
        raise ConversionError(f"PowerPoint 转换 PPTX 失败：{e}", needs_office=True)


def _ppt_extract_text(src: Path) -> list:
    _require_office()
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "out.pptx"
        import win32com.client
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                pres = app.Presentations.Open(str(src.resolve()), ReadOnly=True, WithWindow=False)
                try:
                    pres.SaveAs(str(tmp), 24)
                finally:
                    pres.Close()
            finally:
                app.Quit()
        except Exception as e:
            raise ConversionError(f"PowerPoint 转换失败：{e}", needs_office=True)
        return _pptx_extract(tmp)


@register("ppt", "txt")
def ppt_to_txt(src: Path, dst: Path, options=None):
    blocks = _ppt_extract_text(src)
    dst.write_text("\n".join(t for _, t in blocks), encoding="utf-8")


@register("ppt", "md")
def ppt_to_md_legacy(src: Path, dst: Path, options=None):
    blocks = _ppt_extract_text(src)
    lines = []
    for _, t in blocks:
        lines.append("- " + t)
    dst.write_text("\n".join(lines), encoding="utf-8")
