"""DocForge 通用工具函数。"""

from __future__ import annotations

import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Optional

from charset_normalizer import from_bytes


def detect_encoding(data: bytes) -> str:
    """基于内容嗅探文本编码，兜底 UTF-8。"""
    if not data:
        return "utf-8"
    result = from_bytes(data).best()
    if result and result.encoding:
        return result.encoding
    return "utf-8"


def read_text_auto(path: str | Path) -> str:
    """以自动检测编码的方式读取文本文件。"""
    data = Path(path).read_bytes()
    return data.decode(detect_encoding(data), errors="replace")


def safe_stem(name: str) -> str:
    """清洗文件名主干，去除非法字符。"""
    stem = Path(name).stem
    stem = re.sub(r'[\\/:*?"<>|]', "_", stem).strip()
    stem = unicodedata.normalize("NFC", stem)
    return stem or "output"


def unique_output_path(out_dir: str | Path, stem: str, ext: str) -> Path:
    """生成不冲突的输出路径（重名自动追加序号）。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    ext = ext.lstrip(".").lower()
    candidate = out_dir / f"{stem}.{ext}"
    i = 1
    while candidate.exists():
        candidate = out_dir / f"{stem}_{i}.{ext}"
        i += 1
    return candidate


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def human_size(num_bytes: int) -> str:
    """把字节数格式化为人类可读。"""
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{num_bytes} B"


def timestamp_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def guess_ext_from_name(path: str) -> str:
    """从文件名推断扩展名（归一化）。"""
    return Path(path).suffix.lstrip(".").lower()


def filter_supported(files: list[str], supported_read: set[str]) -> list[str]:
    """过滤出扩展名受支持的输入文件。"""
    out = []
    for f in files:
        ext = guess_ext_from_name(f)
        if ext in supported_read:
            out.append(f)
    return out
