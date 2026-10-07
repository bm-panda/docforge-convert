"""DocForge 命令行接口。

用法示例：
  docforge convert report.docx -t pdf
  docforge convert *.md -t html -o ./out
  docforge convert doc.pdf -t png
  docforge formats
  docforge info report.pdf
"""

from __future__ import annotations

import argparse
import glob as globmod
import sys
from pathlib import Path

from docforge import APP_NAME, __version__
from docforge.core.engine import DocForgeEngine
from docforge.core.models import FORMAT_CATALOG, normalize_ext


def _expand_paths(patterns: list[str]) -> list[str]:
    files = []
    for p in patterns:
        if any(ch in p for ch in "*?["):
            files.extend(globmod.glob(p, recursive=True))
        else:
            files.append(p)
    # 只保留文件
    return [f for f in files if Path(f).is_file()]


def _print_formats():
    eng = DocForgeEngine()
    supported = eng.list_supported()
    print(f"{APP_NAME} v{__version__}\n支持的转换矩阵：\n")
    for src in sorted(supported):
        info = FORMAT_CATALOG.get(src)
        cat = info.category if info else "?"
        dsts = ", ".join(supported[src])
        print(f"  [{cat}] .{src:<5} →  {dsts}")


def _cmd_convert(args) -> int:
    eng = DocForgeEngine()
    files = _expand_paths(args.files)
    if not files:
        print("错误：没有匹配到任何文件", file=sys.stderr)
        return 2

    dst_fmt = normalize_ext(args.to)
    if not FORMAT_CATALOG.get(dst_fmt):
        print(f"错误：不支持的目标格式 .{dst_fmt}", file=sys.stderr)
        return 2

    out_dir = Path(args.out) if args.out else None

    report = eng.convert_batch(files, dst_fmt, out_dir=out_dir, progress_cb=lambda i, t: None)
    for r in report.results:
        tag = {"success": "OK", "failed": "FAIL", "skipped": "SKIP"}[r.status.value]
        print(f"[{tag}] {r.source} -> {r.target or '(无)'}  {r.message}")
    print(report.summary())
    return 0 if report.failed == 0 else 1


def _cmd_info(args) -> int:
    eng = DocForgeEngine()
    for f in args.files:
        p = Path(f)
        if not p.exists():
            print(f"{f}: 文件不存在")
            continue
        fmt = eng.detect_format(p)
        info = eng.get_format_info(fmt)
        readable = "可读取" if fmt in eng.readable_formats else "不可读取"
        desc = info.description if info else "未知格式"
        print(f"{f}\n  格式: .{fmt}（{desc}）  状态: {readable}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docforge",
        description=f"{APP_NAME} —— 全能文档格式转换软件 (v{__version__})",
    )
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {__version__}")
    sub = parser.add_subparsers(dest="command")

    p_conv = sub.add_parser("convert", help="转换一个或多个文件")
    p_conv.add_argument("files", nargs="+", help="输入文件或通配符")
    p_conv.add_argument("-t", "--to", required=True, help="目标格式，如 pdf/docx/md")
    p_conv.add_argument("-o", "--out", help="输出目录（默认输出到源文件目录）")

    sub.add_parser("formats", help="列出全部支持的转换矩阵")

    p_info = sub.add_parser("info", help="检测文件格式")
    p_info.add_argument("files", nargs="+", help="文件路径")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "convert":
        return _cmd_convert(args)
    if args.command == "formats":
        _print_formats()
        return 0
    if args.command == "info":
        return _cmd_info(args)
    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
