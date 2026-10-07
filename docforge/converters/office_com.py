"""基于 MS Office COM 的高保真转换（可选增强路径）。

仅在 Windows 且安装了对应 Office 组件时可用；任何失败都会返回 False，
由调用方回退到纯 Python 方案。本模块绝不抛出异常导致主流程中断。

支持的格式码（Word）：
  wdFormatPDF=17, wdFormatRTF=6, wdFormatText=2, wdFormatUnicodeText=7,
  wdFormatHTML=8, wdFormatXMLDocument=12(docx), wdFormatDocument=0(doc)
PowerPoint: ppSaveAsPDF=32, ppSaveAsPNG=18
Excel: xlFormatPDF=0 / xlTypePDF=0
"""

from __future__ import annotations

import os
from pathlib import Path


def _have_com() -> bool:
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def _genuine(progid: str) -> bool:
    """避免打开用户正在使用的文档，使用独立实例。"""
    return False


def word_to_pdf(src: Path, dst: Path) -> bool:
    """Word 文档 → PDF（高保真）。支持 doc/docx/rtf/odt。"""
    if not _have_com():
        return False
    import win32com.client
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(str(src.resolve()), ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(
                    str(dst.resolve()), 17, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                )
                return Path(dst).exists() and Path(dst).stat().st_size > 0
            finally:
                doc.Close(False)
        finally:
            word.Quit()
    except Exception:
        return False


def word_save_as(src: Path, dst: Path, fmt: int) -> bool:
    """用 Word 打开源文件另存为目标格式。fmt 为 Word 格式码。"""
    if not _have_com():
        return False
    import win32com.client
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(str(src.resolve()), ReadOnly=True)
            try:
                doc.SaveAs2(str(dst.resolve()), FileFormat=fmt)
                return Path(dst).exists() and Path(dst).stat().st_size > 0
            finally:
                doc.Close(False)
        finally:
            word.Quit()
    except Exception:
        return False


def ppt_to_pdf(src: Path, dst: Path) -> bool:
    """PowerPoint 演示文稿 → PDF（高保真）。支持 ppt/pptx。"""
    if not _have_com():
        return False
    import win32com.client
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            pres = app.Presentations.Open(str(src.resolve()), ReadOnly=True, WithWindow=False)
            try:
                pres.SaveAs(str(dst.resolve()), 32)  # ppSaveAsPDF
                return Path(dst).exists() and Path(dst).stat().st_size > 0
            finally:
                pres.Close()
        finally:
            app.Quit()
    except Exception:
        return False


def excel_to_pdf(src: Path, dst: Path) -> bool:
    """Excel 工作簿 → PDF（高保真）。支持 xls/xlsx/ods。"""
    if not _have_com():
        return False
    import win32com.client
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            wb = app.Workbooks.Open(str(src.resolve()), ReadOnly=True)
            try:
                wb.ExportAsFixedFormat(0, str(dst.resolve()))  # xlTypePDF
                return Path(dst).exists() and Path(dst).stat().st_size > 0
            finally:
                wb.Close(False)
        finally:
            app.Quit()
    except Exception:
        return False
