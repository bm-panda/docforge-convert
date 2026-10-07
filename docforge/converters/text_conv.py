"""文本/标记类格式转换：TXT、MD、HTML、RTF、JSON、YAML、XML。

注册的转换对：
  txt -> pdf/docx/html/md/rtf/json/yaml
  md  -> html/pdf/docx/txt/rtf
  html-> md/txt/pdf/docx/rtf
  rtf -> txt/docx/md/html/pdf
  json-> txt/md/html/csv/yaml
  yaml-> json/txt/md/html
  xml -> json/txt/md/html
"""

from __future__ import annotations

import html as html_lib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from docforge.converters.base import ConversionError, register
from docforge.converters.html_pdf import html_to_pdf, text_to_pdf
from docforge.core.utils import read_text_auto, safe_stem

try:
    import markdown as md_lib
    _MD_OK = True
except Exception:
    _MD_OK = False

try:
    import html2text
    _H2T_OK = True
except Exception:
    _H2T_OK = False

try:
    from striprtf.striprtf import rtf_to_text
    _STRIPRTF_OK = True
except Exception:
    _STRIPRTF_OK = False

try:
    import yaml
    _YAML_OK = True
except Exception:
    _YAML_OK = False


# ---------------------------------------------------------------------------
# 公共工具
# ---------------------------------------------------------------------------

def _md_to_html(text: str) -> str:
    if not _MD_OK:
        raise ConversionError("缺少 markdown 库，无法转换 Markdown")
    return md_lib.markdown(
        text,
        extensions=["extra", "codehilite", "tables", "fenced_code", "sane_lists"],
    )


def _html_to_text(html_text: str, as_markdown: bool = False) -> str:
    if _H2T_OK:
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.body_width = 0
        h.ignore_images = True
        h.mark_code = False
        return h.handle(html_text)
    # 兜底：去除标签
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", html_text)).strip()


def _text_to_md(text: str) -> str:
    """纯文本转 Markdown：按需转义，保留换行。"""
    lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            lines.append("")
        elif stripped.startswith(("#", "-", "*", ">", "|", "```")):
            lines.append("\\" + stripped)  # 转义避免被误解析
        else:
            lines.append(stripped)
    return "\n".join(lines)


def _rtf_escape(text: str) -> str:
    out = []
    for ch in text:
        o = ord(ch)
        if ch in "\\{}":
            out.append("\\" + ch)
        elif o < 128:
            if ch == "\n":
                out.append("\\par ")
            else:
                out.append(ch)
        else:
            n = o if o <= 32767 else o - 65536
            out.append(f"\\u{n}?")
    return "".join(out)


def _text_to_rtf(text: str) -> str:
    body = _rtf_escape(text)
    return (
        r"{\rtf1\ansi\ansicpg936\deff0"
        r"{\fonttbl{\f0\fnil\fcharset134 SimSun;}}"
        r"\viewkind4\uc1\pard\lang2052\f0\fs24 "
        + body +
        r"\par}"
    )


def _strip_rtf(text: str) -> str:
    if _STRIPRTF_OK:
        try:
            return rtf_to_text(text, errors="ignore")
        except Exception:
            pass
    # 兜底：去掉 RTF 控制字与分组
    text = re.sub(r"\\u-?\d+ ?", "", text)
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", text)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", text)
    text = re.sub(r"[{}]", "", text)
    return text


def _yaml_dump(data) -> str:
    if not _YAML_OK:
        raise ConversionError("缺少 PyYAML 库，无法转换 YAML")
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False, default_flow_style=False)


def _yaml_load(text: str):
    if not _YAML_OK:
        raise ConversionError("缺少 PyYAML 库，无法转换 YAML")
    return yaml.safe_load(text)


# ---------------------------------------------------------------------------
# DOCX 写入工具
# ---------------------------------------------------------------------------

def _set_cjk_font(run, name: str = "宋体"):
    run.font.name = name
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name)


class _HTMLToDocx:
    """把 HTML 转成 python-docx 文档（块级 + 行内粗斜体/代码）。"""

    INLINE = {"b": ("b",), "strong": ("b",), "i": ("i",), "em": ("i",),
              "u": ("u",), "code": ("code",), "a": ("link",)}

    def __init__(self):
        self.doc = Document()
        style = self.doc.styles["Normal"]
        style.font.name = "Calibri"
        style.font.size = Pt(10.5)
        self._stack: list[str] = []
        self._pending_par = None      # 正在累积的段落
        self._runs: list[tuple[str, set]] = []
        self._list_stack: list[str] = []
        self._in_pre = False
        self._pre_lines: list[str] = []

    def feed(self, html_text: str):
        import html.parser
        p = _HtmlParser(self)
        p.feed(html_text)
        p.close()
        self._flush_paragraph()

    # --- 内部状态 ---
    def _flush_paragraph(self):
        if self._in_pre:
            return
        if self._runs or self._pending_par:
            if self._pending_par is None:
                self._pending_par = self.doc.add_paragraph()
            for text, fmt in self._runs:
                run = self._pending_par.add_run(text)
                _set_cjk_font(run)
                if "b" in fmt:
                    run.bold = True
                if "i" in fmt:
                    run.italic = True
                if "u" in fmt:
                    run.underline = True
                if "code" in fmt:
                    run.font.name = "Consolas"
                    run.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)
            self._runs = []
            self._pending_par = None

    def open_block(self, tag):
        self._flush_paragraph()
        if tag in ("ul", "ol"):
            self._list_stack.append(tag)
        elif tag == "li":
            self._pending_par = self.doc.add_paragraph(style="List Bullet" if self._list_stack and self._list_stack[-1] == "ul" else "List Number")
        elif tag == "pre":
            self._in_pre = True
            self._pre_lines = []
        elif tag == "h1":
            self._pending_par = self.doc.add_heading("", level=1)
        elif tag == "h2":
            self._pending_par = self.doc.add_heading("", level=2)
        elif tag == "h3":
            self._pending_par = self.doc.add_heading("", level=3)
        elif tag == "h4":
            self._pending_par = self.doc.add_heading("", level=4)
        elif tag == "h5":
            self._pending_par = self.doc.add_heading("", level=5)
        elif tag == "h6":
            self._pending_par = self.doc.add_heading("", level=6)
        elif tag == "blockquote":
            self._pending_par = self.doc.add_paragraph()
            self._pending_par.paragraph_format.left_indent = Pt(24)
        elif tag == "hr":
            p = self.doc.add_paragraph()
            run = p.add_run("—" * 20)
            run.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
            self._pending_par = None
        else:
            self._pending_par = self.doc.add_paragraph()

    def close_block(self, tag):
        if tag in ("ul", "ol"):
            if self._list_stack:
                self._list_stack.pop()
        elif tag == "pre":
            self._in_pre = False
            self._pending_par = self.doc.add_paragraph()
            self._pending_par.paragraph_format.left_indent = Pt(18)
            run = self._pending_par.add_run("\n".join(self._pre_lines))
            run.font.name = "Consolas"
            run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
            self._runs = []
            self._pending_par = None
        elif tag == "table":
            self._emit_table()
        else:
            self._flush_paragraph()

    def add_text(self, text: str):
        if self._in_pre:
            self._pre_lines.append(text)
            return
        fmt = set()
        for t in reversed(self._stack):
            if t in ("b", "strong"):
                fmt.add("b")
            elif t in ("i", "em"):
                fmt.add("i")
            elif t == "u":
                fmt.add("u")
            elif t == "code":
                fmt.add("code")
        if text:
            self._runs.append((text, fmt))

    # 表格支持（简化为文本行）
    def _emit_table(self):
        pass


class _HtmlParser(html_lib.parser.HTMLParser):
    def __init__(self, writer):
        super().__init__(convert_charrefs=True)
        self.w = writer

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "br":
            self.w.add_text("\n")
            return
        if tag == "p" and not self.w._stack:
            self.w.open_block("p")
            self.w._stack.append("p")
            return
        if tag in ("ul", "ol", "li", "pre", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "hr"):
            self.w.open_block(tag)
            if tag not in ("hr",):
                self.w._stack.append(tag)
            return
        if tag in ("b", "strong", "i", "em", "u", "code"):
            self.w._stack.append(tag)
            return
        if tag == "table":
            return

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("b", "strong", "i", "em", "u", "code"):
            if tag in self.w._stack:
                self.w._stack.remove(tag)
            return
        if tag in ("ul", "ol", "li", "pre", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "p", "hr"):
            if tag in self.w._stack:
                self.w._stack.remove(tag)
            self.w.close_block(tag)

    def handle_data(self, data):
        self.w.add_text(data)


def _markdown_to_docx(md_text: str, dst: str | Path) -> None:
    html_text = _md_to_html(md_text)
    _html_to_docx(html_text, dst)


def _html_to_docx(html_text: str, dst: str | Path) -> None:
    writer = _HTMLToDocx()
    writer.feed(html_text)
    writer.doc.save(str(dst))


def _text_to_docx(text: str, dst: str | Path) -> None:
    doc = Document()
    for line in text.splitlines():
        if line.strip():
            p = doc.add_paragraph()
            run = p.add_run(line)
            _set_cjk_font(run)
    doc.save(str(dst))


def _json_to_md(data) -> str:
    if isinstance(data, list):
        lines = []
        for i, item in enumerate(data, 1):
            if isinstance(item, dict):
                lines.append(f"### 记录 {i}\n")
                for k, v in item.items():
                    lines.append(f"- **{k}**：{v}")
            else:
                lines.append(f"- {item}")
        return "\n".join(lines)
    lines = []
    for k, v in data.items():
        if isinstance(v, (dict, list)):
            lines.append(f"## {k}\n\n```\n{json.dumps(v, ensure_ascii=False, indent=2)}\n```")
        else:
            lines.append(f"- **{k}**：{v}")
    return "\n".join(lines)


def _json_to_html(data) -> str:
    body = json.dumps(data, ensure_ascii=False, indent=2)
    return (
        "<html><head><meta charset='utf-8'><title>JSON 文档</title></head>"
        f"<body><pre>{html_lib.escape(body)}</pre></body></html>"
    )


# ---------------------------------------------------------------------------
# TXT 目标转换
# ---------------------------------------------------------------------------

@register("txt", "pdf")
def txt_to_pdf(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    text_to_pdf(text, dst, title=src.stem)


@register("txt", "docx")
def txt_to_docx(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    _text_to_docx(text, dst)


@register("txt", "html")
def txt_to_html(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    body = "".join(
        f"<p>{html_lib.escape(line)}</p>" if line.strip() else "<p>&nbsp;</p>"
        for line in text.splitlines()
    )
    dst.write_text(
        "<html><head><meta charset='utf-8'><title>{}</title></head>"
        "<body style='font-family:SimSun,serif;line-height:1.6'>{}</body></html>"
        .format(html_lib.escape(src.stem), body),
        encoding="utf-8",
    )


@register("txt", "md")
def txt_to_md(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    dst.write_text(_text_to_md(text), encoding="utf-8")


@register("txt", "rtf")
def txt_to_rtf(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    dst.write_text(_text_to_rtf(text), encoding="utf-8")


@register("txt", "json")
def txt_to_json(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    data = {"source": src.stem, "lines": text.splitlines(), "content": text}
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register("txt", "yaml")
def txt_to_yaml(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    data = {"source": src.stem, "lines": text.splitlines(), "content": text}
    dst.write_text(_yaml_dump(data), encoding="utf-8")


# ---------------------------------------------------------------------------
# Markdown 目标转换
# ---------------------------------------------------------------------------

@register("md", "html")
def md_to_html(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    body = _md_to_html(text)
    full = (
        "<html><head><meta charset='utf-8'><title>{}</title></head>"
        "<body style='max-width:820px;margin:0 auto;padding:24px;"
        "font-family:SimSun,serif;line-height:1.7'>{}</body></html>"
        .format(html_lib.escape(src.stem), body)
    )
    dst.write_text(full, encoding="utf-8")


@register("md", "pdf")
def md_to_pdf(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    html_text = _md_to_html(text)
    html_to_pdf(html_text, dst, title=src.stem)


@register("md", "docx")
def md_to_docx(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    _markdown_to_docx(text, dst)


@register("md", "txt")
def md_to_txt(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    html_text = _md_to_html(text)
    dst.write_text(_html_to_text(html_text), encoding="utf-8")


@register("md", "rtf")
def md_to_rtf(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    html_text = _md_to_html(text)
    plain = _html_to_text(html_text)
    dst.write_text(_text_to_rtf(plain), encoding="utf-8")


# ---------------------------------------------------------------------------
# HTML 目标转换
# ---------------------------------------------------------------------------

@register("html", "md")
def html_to_md(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    dst.write_text(_html_to_text(text, as_markdown=True), encoding="utf-8")


@register("html", "txt")
def html_to_txt(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    dst.write_text(_html_to_text(text), encoding="utf-8")


@register("html", "pdf")
def html_to_pdf_conv(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    html_to_pdf(text, dst, title=src.stem)


@register("html", "docx")
def html_to_docx_conv(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    _html_to_docx(text, dst)


@register("html", "rtf")
def html_to_rtf_conv(src: Path, dst: Path, options=None):
    text = read_text_auto(src)
    plain = _html_to_text(text)
    dst.write_text(_text_to_rtf(plain), encoding="utf-8")


# ---------------------------------------------------------------------------
# RTF 目标转换
# ---------------------------------------------------------------------------

@register("rtf", "txt")
def rtf_to_txt(src: Path, dst: Path, options=None):
    data = src.read_bytes()
    text = data.decode("utf-8", errors="replace")
    dst.write_text(_strip_rtf(text), encoding="utf-8")


def _rtf_to_plain(src: Path) -> str:
    data = src.read_bytes()
    text = data.decode("utf-8", errors="replace")
    return _strip_rtf(text)


@register("rtf", "docx")
def rtf_to_docx(src: Path, dst: Path, options=None):
    plain = _rtf_to_plain(src)
    _text_to_docx(plain, dst)


@register("rtf", "md")
def rtf_to_md(src: Path, dst: Path, options=None):
    plain = _rtf_to_plain(src)
    dst.write_text(_text_to_md(plain), encoding="utf-8")


@register("rtf", "html")
def rtf_to_html(src: Path, dst: Path, options=None):
    plain = _rtf_to_plain(src)
    body = "".join(
        f"<p>{html_lib.escape(line)}</p>" if line.strip() else "<p>&nbsp;</p>"
        for line in plain.splitlines()
    )
    dst.write_text(
        "<html><head><meta charset='utf-8'><title>{}</title></head>"
        "<body style='font-family:SimSun,serif;line-height:1.6'>{}</body></html>"
        .format(html_lib.escape(src.stem), body),
        encoding="utf-8",
    )


@register("rtf", "pdf")
def rtf_to_pdf(src: Path, dst: Path, options=None):
    plain = _rtf_to_plain(src)
    text_to_pdf(plain, dst, title=src.stem)


# ---------------------------------------------------------------------------
# JSON 目标转换
# ---------------------------------------------------------------------------

def _read_json(src: Path):
    text = read_text_auto(src)
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise ConversionError(f"JSON 解析失败：{e}")


@register("json", "txt")
def json_to_txt(src: Path, dst: Path, options=None):
    data = _read_json(src)
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register("json", "md")
def json_to_md_conv(src: Path, dst: Path, options=None):
    data = _read_json(src)
    dst.write_text(_json_to_md(data), encoding="utf-8")


@register("json", "html")
def json_to_html_conv(src: Path, dst: Path, options=None):
    data = _read_json(src)
    dst.write_text(_json_to_html(data), encoding="utf-8")


@register("json", "csv")
def json_to_csv(src: Path, dst: Path, options=None):
    import csv
    data = _read_json(src)
    records = data if isinstance(data, list) else data.get("records") or data.get("data") or []
    if not isinstance(records, list) or not records:
        raise ConversionError("JSON 顶层不是对象数组，无法转为 CSV")
    if not all(isinstance(r, dict) for r in records):
        raise ConversionError("JSON 数组元素不是对象，无法转为 CSV")
    keys = list(records[0].keys())
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        for r in records:
            w.writerow({k: "" if r.get(k) is None else r[k] for k in keys})


@register("json", "yaml")
def json_to_yaml(src: Path, dst: Path, options=None):
    data = _read_json(src)
    dst.write_text(_yaml_dump(data), encoding="utf-8")


# ---------------------------------------------------------------------------
# YAML 目标转换
# ---------------------------------------------------------------------------

def _read_yaml(src: Path):
    text = read_text_auto(src)
    data = _yaml_load(text)
    if data is None:
        data = {}
    return data


@register("yaml", "json")
def yaml_to_json(src: Path, dst: Path, options=None):
    data = _read_yaml(src)
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register("yaml", "txt")
def yaml_to_txt(src: Path, dst: Path, options=None):
    data = _read_yaml(src)
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register("yaml", "md")
def yaml_to_md(src: Path, dst: Path, options=None):
    data = _read_yaml(src)
    dst.write_text(_json_to_md(data), encoding="utf-8")


@register("yaml", "html")
def yaml_to_html(src: Path, dst: Path, options=None):
    data = _read_yaml(src)
    dst.write_text(_json_to_html(data), encoding="utf-8")


# ---------------------------------------------------------------------------
# XML 目标转换
# ---------------------------------------------------------------------------

def _xml_to_dict(elem: ET.Element) -> dict:
    def conv(el):
        children = list(el)
        if not children:
            return el.text or ""
        if all(c.tag == el.tag + "_item" for c in children):
            return [conv(c) for c in children]
        d = {}
        for c in children:
            key = c.tag.split("}")[-1]
            if key in d:
                if not isinstance(d[key], list):
                    d[key] = [d[key]]
                d[key].append(conv(c))
            else:
                d[key] = conv(c)
        return d
    return {elem.tag.split("}")[-1]: conv(elem)}


def _read_xml(src: Path):
    try:
        tree = ET.parse(src)
    except ET.ParseError as e:
        raise ConversionError(f"XML 解析失败：{e}")
    return _xml_to_dict(tree.getroot())


@register("xml", "json")
def xml_to_json(src: Path, dst: Path, options=None):
    data = _read_xml(src)
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register("xml", "txt")
def xml_to_txt(src: Path, dst: Path, options=None):
    data = _read_xml(src)
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@register("xml", "md")
def xml_to_md(src: Path, dst: Path, options=None):
    data = _read_xml(src)
    dst.write_text(_json_to_md(data), encoding="utf-8")


@register("xml", "html")
def xml_to_html(src: Path, dst: Path, options=None):
    data = _read_xml(src)
    dst.write_text(_json_to_html(data), encoding="utf-8")
