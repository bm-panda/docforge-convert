"""DocForge 核心引擎：格式检测、单文件/批量转换调度。"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Callable, Optional

from docforge.converters import get_converter, supported_destinations
from docforge.converters.base import ConversionError, ConversionOptions
from docforge.core.models import (
    FORMAT_CATALOG,
    BatchReport,
    ConversionResult,
    ConversionStatus,
    FormatInfo,
    normalize_ext,
)
from docforge.core.utils import filter_supported, guess_ext_from_name, safe_stem, unique_output_path

# 目录型输出格式：转换后是一个文件夹（图片序列等）
DIRECTORY_OUTPUT = {"png", "jpg"}


class DocForgeEngine:
    """文档转换引擎。线程安全（每次调用互不共享可变状态）。"""

    def __init__(self):
        self._readable_cache: Optional[set] = None

    # ------------------------------------------------------------------
    # 格式信息
    # ------------------------------------------------------------------
    @property
    def readable_formats(self) -> set:
        if self._readable_cache is None:
            from docforge.converters import CAPABILITIES
            self._readable_cache = set(CAPABILITIES.keys())
        return self._readable_cache

    def get_format_info(self, ext: str) -> Optional[FormatInfo]:
        return FORMAT_CATALOG.get(normalize_ext(ext))

    def detect_format(self, path: str | Path) -> str:
        """根据文件扩展名识别格式（归一化）。"""
        return normalize_ext(guess_ext_from_name(str(path)))

    def list_supported(self) -> dict[str, list[str]]:
        """返回 {源格式: [目标格式...]}（按类别排序）。"""
        result = {}
        for src in sorted(self.readable_formats):
            dsts = sorted(supported_destinations(src))
            if dsts:
                result[src] = dsts
        return result

    def is_readable(self, path: str | Path) -> bool:
        return self.detect_format(path) in self.readable_formats

    # ------------------------------------------------------------------
    # 转换
    # ------------------------------------------------------------------
    def convert_one(self, src: str | Path, dst_fmt: str,
                    out_dir: Optional[str | Path] = None,
                    options: Optional[dict] = None) -> ConversionResult:
        """转换单个文件。out_dir 缺省时输出到源文件目录。"""
        src = Path(src)
        dst_fmt = normalize_ext(dst_fmt)
        src_fmt = self.detect_format(src)
        options = ConversionOptions(options or {})

        start = time.perf_counter()
        target = ""
        try:
            if not src.exists():
                raise ConversionError("源文件不存在")
            if src_fmt == dst_fmt:
                return ConversionResult(
                    str(src), str(src), ConversionStatus.SKIPPED,
                    "源格式与目标格式相同", 0.0, src_fmt, dst_fmt,
                )
            if src_fmt not in self.readable_formats:
                raise ConversionError(f"暂不支持读取 {src_fmt.upper()} 格式")

            converter = get_converter(src_fmt, dst_fmt)
            if converter is None:
                raise ConversionError(
                    f"暂不支持 {src_fmt.upper()} → {dst_fmt.upper()} 转换"
                )

            out_dir = Path(out_dir) if out_dir else src.parent
            if dst_fmt in DIRECTORY_OUTPUT:
                # 目录型输出：例如 PDF→PNG 多页
                stem = safe_stem(src.name)
                folder = out_dir / f"{stem}_to_{dst_fmt}"
                folder.mkdir(parents=True, exist_ok=True)
                target = str(folder / f"{stem}_p1.{dst_fmt}")
                dst = Path(target)
            else:
                target = str(unique_output_path(out_dir, safe_stem(src.name), dst_fmt))
                dst = Path(target)

            converter(src, dst, options)

            if not dst.exists():
                raise ConversionError("转换完成但未生成输出文件")

            dur = time.perf_counter() - start
            return ConversionResult(
                str(src), str(dst), ConversionStatus.SUCCESS,
                "转换成功", dur, src_fmt, dst_fmt,
            )
        except ConversionError as e:
            dur = time.perf_counter() - start
            return ConversionResult(
                str(src), target, ConversionStatus.FAILED,
                e.message, dur, src_fmt, dst_fmt,
            )
        except Exception as e:  # noqa: BLE001
            dur = time.perf_counter() - start
            return ConversionResult(
                str(src), target, ConversionStatus.FAILED,
                f"转换异常：{e}", dur, src_fmt, dst_fmt,
            )

    def convert_batch(self, files: list[str | Path], dst_fmt: str,
                      out_dir: Optional[str | Path] = None,
                      options: Optional[dict] = None,
                      progress_cb: Optional[Callable[[int, int], None]] = None,
                      ) -> BatchReport:
        """批量转换。progress_cb(i, total) 在每次完成后回调。"""
        report = BatchReport(start_time=time.perf_counter())
        dst_fmt = normalize_ext(dst_fmt)
        supported = self.readable_formats
        files = [str(f) for f in files]

        candidates = filter_supported(files, supported)
        unsupported = [f for f in files if f not in candidates]
        for f in unsupported:
            report.results.append(ConversionResult(
                f, "", ConversionStatus.FAILED,
                f"暂不支持该格式读取", 0.0, self.detect_format(f), dst_fmt,
            ))
            report.failed += 1
        report.total = len(files)

        total_done = 0
        for idx, f in enumerate(candidates):
            res = self.convert_one(f, dst_fmt, out_dir, options)
            report.results.append(res)
            if res.status == ConversionStatus.SUCCESS:
                report.success += 1
            elif res.status == ConversionStatus.FAILED:
                report.failed += 1
            else:
                report.skipped += 1
            total_done += 1
            if progress_cb:
                progress_cb(total_done, len(files))

        report.end_time = time.perf_counter()
        return report


# 单例便于直接使用
engine = DocForgeEngine()
