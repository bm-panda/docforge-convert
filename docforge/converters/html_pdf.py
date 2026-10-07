"""轻量 HTML → PDF 渲染器（reportlab Platypus 实现，完全离线）。

支持块级：h1-h6 / p / ul / ol / li / table / pre / blockquote / hr / br / div
支持行内：b / strong / i / em / u / code / a / span
中文字体使用内建 CID 字体 STSong-Light，无需外部字体文件。
"""

from __future__ import annotations

import html
import re
from html.parser import HTMLParser
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

try:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    _FONT = "STSong-Light"
except Exception:
    _FONT = "Helvetica"


class _Block:
    """一个块级元素（最终产出为一个或多个 Flowable）。"""

    __slots__ = ("kind", "lines", "attrs")

    def __init__(self, kind: str, attrs: dict):
        self.kind = kind
        self.lines: list[str] = []
        self.attrs = attrs


_INLINE_MAP = {
    "b": "b", "strong": "b", "i": "i", "em": "i",
    "u": "u", "code": "font face=monospace", "mark": "font backColor=yellow",
}


def _build_styles() -> dict:
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "df_title", fontName=_FONT, fontSize=22, leading=30,
            alignment=TA_CENTER, spaceAfter=14,
        ),
        "h1": ParagraphStyle("df_h1", fontName=_FONT, fontSize=19, leading=26,
                             spaceBefore=12, spaceAfter=8, textColor=colors.HexColor("#1a1a2e")),
        "h2": ParagraphStyle("df_h2", fontName=_FONT, fontSize=16, leading=22,
                             spaceBefore=10, spaceAfter=6, textColor=colors.HexColor("#16213e")),
        "h3": ParagraphStyle("df_h3", fontName=_FONT, fontSize=14, leading=19,
                             spaceBefore=8, spaceAfter=5),
        "h4": ParagraphStyle("df_h4", fontName=_FONT, fontSize=12.5, leading=17,
                             spaceBefore=6, spaceAfter=4),
        "h5": ParagraphStyle("df_h5", fontName=_FONT, fontSize=11.5, leading=16,
                             spaceBefore=6, spaceAfter=4),
        "h6": ParagraphStyle("df_h6", fontName=_FONT, fontSize=11, leading=15,
                             spaceBefore=5, spaceAfter=3),
        "p": ParagraphStyle("df_p", fontName=_FONT, fontSize=11, leading=17,
                            alignment=TA_JUSTIFY, spaceAfter=6),
        "li": ParagraphStyle("df_li", fontName=_FONT, fontSize=11, leading=17,
                             alignment=TA_JUSTIFY, spaceAfter=3),
        "quote": ParagraphStyle("df_quote", fontName=_FONT, fontSize=10.5, leading=16,
                                leftIndent=14, textColor=colors.HexColor("#555555")),
        "table": ParagraphStyle("df_tbl", fontName=_FONT, fontSize=9.5, leading=13),
        "th": ParagraphStyle("df_th", fontName=_FONT, fontSize=9.5, leading=13),
        "td": ParagraphStyle("df_td", fontName=_FONT, fontSize=9.5, leading=13),
        "code": ParagraphStyle("df_code", fontName="Courier", fontSize=9,
                               leading=13, leftIndent=10,
                               backColor=colors.HexColor("#f5f5f5"),
                               borderPadding=4, spaceAfter=6),
    }
    return styles


class _HTMLFlowableBuilder(HTMLParser):
    """把 HTML 解析为 reportlab Flowable 列表。"""

    BLOCK_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "pre", "blockquote",
                  "div", "ul", "ol", "li", "table", "tr", "td", "th", "hr"}

    def __init__(self, base_size="A4"):
        super().__init__(convert_charrefs=True)
        self.styles = _build_styles()
        self.flowables: list = []
        self._stack: list[_Block] = []
        self._inline_stack: list[str] = []
        self._in_pre = False
        self._table_rows: list[list] = []
        self._cur_row: list = []
        self._cur_cell: list = []

    # ---------- 栈操作 ----------
    def _push(self, kind, attrs):
        self._stack.append(_Block(kind, attrs))

    def _top(self):
        return self._stack[-1] if self._stack else None

    def _flush_inline_text(self):
        """把当前 inline 文本追加到栈顶块。"""
        if not self._inline_stack:
            return
        text = "".join(self._inline_stack).strip()
        self._inline_stack = []
        if not text:
            return
        top = self._top()
        if top is None:
            return
        top.lines.append(text)

    # ---------- 事件 ----------
    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        a = dict(attrs)

        if tag in _INLINE_MAP:
            self._inline_stack.append(f"<{_INLINE_MAP[tag]}>")
            return
        if tag == "a":
            href = a.get("href", "")
            if href:
                self._inline_stack.append(f'<link href="{html.escape(href, quote=True)}" color="blue">')
            else:
                self._inline_stack.append("")
            return
        if tag == "br":
            self._inline_stack.append("<br/>")
            return
        if tag == "img":
            self._inline_stack.append("［图片］")
            return
        if tag == "font":
            color = a.get("color", "")
            self._inline_stack.append(f'<font color="{html.escape(color, quote=True)}">')
            return
        if tag == "hr":
            self._flush_inline_text()
            self.flowables.append(Spacer(1, 6))
            self.flowables.append(Table([[""]], colWidths=[180 * mm],
                                        rowHeights=[1],
                                        style=TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.6, colors.grey)])))
            self.flowables.append(Spacer(1, 6))
            return

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6", "p", "pre", "blockquote", "div"):
            self._flush_inline_text()
            self._push(tag, a)
            self._in_pre = (tag == "pre")
            return

        if tag == "ul" or tag == "ol":
            self._flush_inline_text()
            self._push(tag, a)
            return

        if tag == "li":
            self._flush_inline_text()
            # 结束上一个 li（若有）
            if self._top() and self._top().kind == "li":
                self._pop_block()
            self._push("li", a)
            return

        if tag == "table":
            self._flush_inline_text()
            self._table_rows = []
            return

        if tag == "tr":
            if self._cur_row:
                self._table_rows.append(self._cur_row)
            self._cur_row = []
            return

        if tag in ("td", "th"):
            self._flush_inline_text()
            # 结束上一单元格
            if self._cur_cell:
                self._cur_row.append("".join(self._cur_cell))
                self._cur_cell = []
            self._cur_cell = []
            self._push(tag, a)
            return

    def handle_endtag(self, tag):
        tag = tag.lower()

        if tag in _INLINE_MAP:
            self._inline_stack.append("</b>") if tag in ("b", "strong") else None
            if tag in ("b", "strong"):
                self._inline_stack.append("</b>")
            elif tag in ("i", "em"):
                self._inline_stack.append("</i>")
            elif tag == "u":
                self._inline_stack.append("</u>")
            elif tag == "code":
                self._inline_stack.append("</font>")
            elif tag == "mark":
                self._inline_stack.append("</font>")
            return
        if tag == "a":
            self._inline_stack.append("</link>")
            return
        if tag == "font":
            self._inline_stack.append("</font>")
            return

        if tag in ("td", "th"):
            self._flush_inline_text()
            if self._cur_cell is not None:
                self._cur_row.append("".join(self._cur_cell))
            self._cur_cell = []
            if self._stack and self._stack[-1].kind in ("td", "th"):
                self._stack.pop()
            return

        if tag == "tr":
            if self._cur_row:
                self._table_rows.append(self._cur_row)
                self._cur_row = []
            return

        if tag == "table":
            if self._cur_row:
                self._table_rows.append(self._cur_row)
                self._cur_row = []
            if self._table_rows:
                self._emit_table(self._table_rows)
            self._table_rows = []
            return

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6", "p", "pre", "blockquote", "div", "li", "ul", "ol"):
            self._flush_inline_text()
            self._pop_block()
            if tag == "pre":
                self._in_pre = False
            return

    def handle_data(self, data):
        if self._in_pre:
            top = self._top()
            if top is not None and top.kind == "pre":
                top.lines.append(data)
            return
        if self._top() and self._top().kind in ("td", "th"):
            self._cur_cell.append(data)
            return
        self._inline_stack.append(data)

    # ---------- 块产出 ----------
    def _pop_block(self):
        top = self._top()
        if top is None:
            return
        self._stack.pop()
        if top.kind == "ul":
            self._emit_list(top, ordered=False)
        elif top.kind == "ol":
            self._emit_list(top, ordered=True)
        elif top.kind == "li":
            self._emit_list_item(top)
        else:
            self._emit_simple(top)

    def _emit_simple(self, block):
        text = "\n".join(block.lines).strip()
        if not text and block.kind != "pre":
            return
        style_key = {
            "h1": "h1", "h2": "h2", "h3": "h3", "h4": "h4",
            "h5": "h5", "h6": "h6", "p": "p", "blockquote": "quote",
            "div": "p", "pre": "code",
        }.get(block.kind, "p")
        if block.kind == "pre":
            self.flowables.append(Preformatted(text, self.styles["code"]))
        else:
            self.flowables.append(Paragraph(text, self.styles[style_key]))

    def _emit_list_item(self, block):
        text = " ".join(block.lines).strip()
        if text:
            self.flowables.append(("__list_item__", Paragraph(text, self.styles["li"])))

    def _emit_list(self, block, ordered: bool):
        # 收集紧随的列表项
        items = []
        rest = []
        for fl in self.flowables:
            if isinstance(fl, tuple) and fl[0] == "__list_item__":
                items.append(fl[1])
            else:
                rest.append(fl)
        self.flowables = rest
        if items:
            self.flowables.append(ListFlowable(
                items, bulletType="1" if ordered else "bullet",
                start=None if ordered else "•", leftIndent=16,
            ))

    def _emit_table(self, rows):
        # 去掉最后被吞掉的空行
        while rows and all(not c for c in rows[-1]):
            rows.pop()
        if not rows:
            return
        ncols = max(len(r) for r in rows)
        data = []
        for r in rows:
            row = [Paragraph(c or "", self.styles["td"]) for c in r]
            while len(row) < ncols:
                row.append(Paragraph("", self.styles["td"]))
            data.append(row)
        # 表头样式（第一行）
        style_cmds = [
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef1f6")),
        ]
        tbl = Table(data, repeatRows=1)
        tbl.setStyle(TableStyle(style_cmds))
        self.flowables.append(Spacer(1, 4))
        self.flowables.append(tbl)
        self.flowables.append(Spacer(1, 8))


def html_to_pdf(html_text: str, dst: str | Path, title: str = "DocForge 文档",
                page_size=A4) -> None:
    """将 HTML 字符串渲染为 PDF 文件。"""
    parser = _HTMLFlowableBuilder()
    try:
        parser.feed(html_text)
        parser.close()
    except Exception:
        # 解析失败时退化为纯文本段落
        plain = re.sub(r"<[^>]+>", "", html_text)
        paragraphs = [p.strip() for p in plain.splitlines() if p.strip()]
        parser.flowables = [Paragraph(p, parser.styles["p"]) for p in paragraphs]

    flowables = []
    if title:
        flowables.append(Paragraph(html.escape(title), parser.styles["title"]))
    flowables.extend(f for f in parser.flowables if not (isinstance(f, tuple)))
    flowables.append(Spacer(1, 12))

    doc = SimpleDocTemplate(
        str(dst), pagesize=page_size,
        leftMargin=20 * mm, rightMargin=20 * mm,
        topMargin=18 * mm, bottomMargin=18 * mm,
        title=title, author="DocForge",
    )
    doc.build(flowables)


def text_to_pdf(text: str, dst: str | Path, title: str = "DocForge 文档") -> None:
    """把纯文本渲染为 PDF（自动换行，保留段落）。"""
    safe = html.escape(text)
    paragraphs = [p for p in safe.splitlines()]
    joined = "<br/>".join(paragraphs)
    body = f"<p>{joined}</p>"
    html_to_pdf(body, dst, title=title)
