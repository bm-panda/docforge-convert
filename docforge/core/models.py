"""DocForge 数据模型定义。"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Optional


class ConversionStatus(str, Enum):
    """单次转换的结果状态。"""

    SUCCESS = "success"
    FAILED = "failed"
    SKIPPED = "skipped"


@dataclass
class ConversionResult:
    """单次文件转换的结果。"""

    source: str                      # 源文件绝对路径
    target: str                      # 目标文件绝对路径
    status: ConversionStatus         # 结果状态
    message: str = ""                # 结果描述 / 错误信息
    duration: float = 0.0            # 耗时（秒）
    format_from: str = ""            # 源格式
    format_to: str = ""              # 目标格式

    @property
    def ok(self) -> bool:
        return self.status == ConversionStatus.SUCCESS


@dataclass
class FormatInfo:
    """某一种格式的元信息。"""

    ext: str                         # 扩展名（小写，不含点）
    category: str                    # 类别：文档 / 表格 / 演示 / 图片 / 电子书 / 数据
    description: str = ""            # 中文描述
    readable: bool = True            # 是否可读
    writable: bool = True            # 是否可写


@dataclass
class BatchReport:
    """一次批量转换的汇总报告。"""

    total: int = 0
    success: int = 0
    failed: int = 0
    skipped: int = 0
    results: List[ConversionResult] = field(default_factory=list)
    start_time: float = 0.0
    end_time: float = 0.0

    @property
    def duration(self) -> float:
        return self.end_time - self.start_time

    def summary(self) -> str:
        return (
            f"共 {self.total} 个文件：成功 {self.success}，失败 {self.failed}，"
            f"跳过 {self.skipped}，总耗时 {self.duration:.1f} 秒"
        )


# 格式类别常量
CATEGORY_DOCUMENT = "文档"
CATEGORY_SPREADSHEET = "表格"
CATEGORY_PRESENTATION = "演示"
CATEGORY_IMAGE = "图片"
CATEGORY_EBOOK = "电子书"
CATEGORY_DATA = "数据"


# 全部支持格式的元信息（与转换矩阵配套）
FORMAT_CATALOG: Dict[str, FormatInfo] = {
    "txt":   FormatInfo("txt", CATEGORY_DOCUMENT, "纯文本"),
    "md":    FormatInfo("md", CATEGORY_DOCUMENT, "Markdown"),
    "html":  FormatInfo("html", CATEGORY_DOCUMENT, "网页"),
    "htm":   FormatInfo("htm", CATEGORY_DOCUMENT, "网页(旧)"),
    "rtf":   FormatInfo("rtf", CATEGORY_DOCUMENT, "富文本"),
    "docx":  FormatInfo("docx", CATEGORY_DOCUMENT, "Word 文档"),
    "doc":   FormatInfo("doc", CATEGORY_DOCUMENT, "Word 97 文档"),
    "odt":   FormatInfo("odt", CATEGORY_DOCUMENT, "ODT 文本文档"),
    "pdf":   FormatInfo("pdf", CATEGORY_DOCUMENT, "PDF 文档"),
    "xlsx":  FormatInfo("xlsx", CATEGORY_SPREADSHEET, "Excel 工作簿"),
    "xls":   FormatInfo("xls", CATEGORY_SPREADSHEET, "Excel 97 工作簿"),
    "ods":   FormatInfo("ods", CATEGORY_SPREADSHEET, "ODS 电子表格"),
    "csv":   FormatInfo("csv", CATEGORY_SPREADSHEET, "CSV 表格"),
    "tsv":   FormatInfo("tsv", CATEGORY_SPREADSHEET, "TSV 表格"),
    "pptx":  FormatInfo("pptx", CATEGORY_PRESENTATION, "PPT 演示文稿"),
    "ppt":   FormatInfo("ppt", CATEGORY_PRESENTATION, "PPT 97 演示文稿"),
    "odp":   FormatInfo("odp", CATEGORY_PRESENTATION, "ODP 演示文稿"),
    "png":   FormatInfo("png", CATEGORY_IMAGE, "PNG 图片"),
    "jpg":   FormatInfo("jpg", CATEGORY_IMAGE, "JPEG 图片"),
    "jpeg":  FormatInfo("jpeg", CATEGORY_IMAGE, "JPEG 图片"),
    "epub":  FormatInfo("epub", CATEGORY_EBOOK, "EPUB 电子书"),
    "json":  FormatInfo("json", CATEGORY_DATA, "JSON 数据"),
    "xml":   FormatInfo("xml", CATEGORY_DATA, "XML 数据"),
    "yaml":  FormatInfo("yaml", CATEGORY_DATA, "YAML 数据"),
    "yml":   FormatInfo("yml", CATEGORY_DATA, "YAML 数据"),
}

# 别名归一化：将常见扩展名映射到标准格式名
EXT_ALIASES: Dict[str, str] = {
    "htm": "html",
    "jpeg": "jpg",
    "yml": "yaml",
}


def normalize_ext(ext: str) -> str:
    """将扩展名归一化为标准格式名（去点、小写、合并别名）。"""
    ext = ext.strip().lstrip(".").lower()
    return EXT_ALIASES.get(ext, ext)
