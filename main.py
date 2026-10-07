#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""文档格式转换 —— 不忙脚本盒子入口脚本（Python，无 GUI）。

盒子把参数 JSON 文件路径放在命令行第一个参数（sys.argv[1]）。
本脚本读取盒子参数 -> 调用内置的 DocForge（开源项目，MIT，作者 binghai-fly）完成转换；
转换进度与结果全部通过命令行友好展示（无图形界面），
结束后暂停 5 秒再自动退出；被联动调用时改为回传信封。

内置的 DocForge 包（MIT）位于脚本目录下的 docforge/。本项目是对 DocForge 的
二次封装（Fork），原作者署名与许可见 README 的“开源说明与致谢”。
"""

from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

# 控制台编码兜底：不改编码，仅把无法编码的字符替换掉，避免在 GBK 等代码页下打印崩溃
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

PAUSE_SECONDS = 5


# ---------------------------------------------------------------------------
# 终端友好输出
# ---------------------------------------------------------------------------

class _Style:
    """ANSI 颜色；终端不支持时自动降级为纯文本。"""

    def __init__(self) -> None:
        self.enabled = False
        if os.name == "nt":
            self.enabled = _enable_windows_ansi()
        if hasattr(sys.stdout, "isatty") and not sys.stdout.isatty():
            self.enabled = False
        self.green = "\033[32m" if self.enabled else ""
        self.red = "\033[31m" if self.enabled else ""
        self.yellow = "\033[33m" if self.enabled else ""
        self.cyan = "\033[36m" if self.enabled else ""
        self.bold = "\033[1m" if self.enabled else ""
        self.dim = "\033[2m" if self.enabled else ""
        self.reset = "\033[0m" if self.enabled else ""


def _enable_windows_ansi() -> bool:
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        mode = ctypes.c_uint32()
        if not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            return False
        return bool(kernel32.SetConsoleMode(handle, mode.value | 0x0004))
    except Exception:
        return False


STYLE = _Style()


def _hr(char: str = "-", width: int = 52) -> str:
    return char * width


def _progress_bar(done: int, total: int, width: int = 18) -> str:
    if total <= 0:
        return "[" + " " * width + "]"
    filled = int(width * done / total)
    return "[" + "#" * filled + "-" * (width - filled) + f"] {done}/{total}"


def _fmt_duration(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, sec = divmod(seconds, 60)
    return f"{int(minutes)}m{sec:.0f}s"


# ---------------------------------------------------------------------------
# 参数读取与信封
# ---------------------------------------------------------------------------

def read_payload() -> dict:
    if len(sys.argv) < 2:
        raise RuntimeError("缺少参数 JSON 路径（argv[1]）")
    with open(sys.argv[1], "r", encoding="utf-8") as handle:
        return json.load(handle)


def write_envelope(env: dict, payload: dict) -> None:
    output_json = (env or {}).get("output_json")
    if not output_json:
        return
    with open(output_json, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def normalize_files(value) -> list:
    if not isinstance(value, list):
        value = [value] if value else []
    return [str(item).strip() for item in value if str(item or "").strip()]


def collect_outputs(result) -> list:
    """返回本次结果实际产出的文件路径（目录型输出展开为多页）。"""
    path = Path(result.target)
    if path.is_dir():
        return [str(p) for p in sorted(path.rglob("*")) if p.is_file()]
    if path.is_file():
        parent = path.parent
        if parent.name.endswith(f"_to_{result.format_to}") and parent.is_dir():
            return [str(p) for p in sorted(parent.glob(f"*{path.suffix}"))]
        return [str(path)]
    return []


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main() -> None:
    try:
        payload = read_payload()
    except Exception as exc:  # noqa: BLE001
        print(f"读取参数失败：{exc}", file=sys.stderr)
        sys.exit(1)

    env = payload.get("environment") or {}
    data = payload.get("data") or {}
    params = payload.get("params") or {}
    node_call = env.get("invoke_mode") == "node"

    try:
        from docforge.core.engine import DocForgeEngine
        from docforge.core.models import ConversionStatus, normalize_ext
    except Exception as exc:  # noqa: BLE001
        message = f"转换引擎加载失败（依赖未安装？）：{exc}"
        if node_call:
            write_envelope(env, {"code": 1, "msg": message})
        else:
            print(f"转换失败：{message}", file=sys.stderr)
            _pause()
        sys.exit(1)

    files = normalize_files(data.get("input_files"))
    target = normalize_ext(str(params.get("target") or "pdf"))
    output_dir = str(params.get("output_dir") or "").strip()

    if not files:
        message = "未选择源文件"
        if node_call:
            write_envelope(env, {"code": 1, "msg": message})
        else:
            print(f"转换失败：{message}", file=sys.stderr)
            _pause()
        sys.exit(1)

    engine = DocForgeEngine()
    supported = set(engine.readable_formats)
    out_dir = Path(output_dir) if output_dir else None

    _print_header(files, target, out_dir)

    records = []
    started = time.perf_counter()
    succeeded = failed = skipped = 0

    for index, item in enumerate(files, 1):
        name = Path(item).name
        prefix = f"[{index}/{len(files)}]"
        fmt = engine.detect_format(item)

        if fmt not in supported:
            message = "暂不支持该格式读取"
            print(f"{prefix} {STYLE.red}跳过{STYLE.reset} {name}  {STYLE.dim}{message}{STYLE.reset}")
            records.append(_record(item, "", ConversionStatus.FAILED.value, message, []))
            failed += 1
            continue

        print(f"{prefix} {name} {STYLE.dim}->{STYLE.reset} {target.upper()} ", end="", flush=True)
        result = engine.convert_one(item, target, out_dir=out_dir)
        outputs = collect_outputs(result) if result.status == ConversionStatus.SUCCESS else []
        records.append(_record(result.source, result.target, result.status.value, result.message, outputs))

        if result.status == ConversionStatus.SUCCESS:
            succeeded += 1
            detail = _output_label(outputs)
            print(f"{STYLE.green}完成{STYLE.reset} {STYLE.dim}{result.duration:.1f}s{STYLE.reset}  {detail}")
        elif result.status == ConversionStatus.SKIPPED:
            skipped += 1
            print(f"{STYLE.yellow}跳过{STYLE.reset} {STYLE.dim}{result.message}{STYLE.reset}")
        else:
            failed += 1
            print(f"{STYLE.red}失败{STYLE.reset} {STYLE.dim}{result.message}{STYLE.reset}")

    elapsed = time.perf_counter() - started
    _print_summary(succeeded, failed, skipped, len(files), elapsed)

    if node_call:
        ok = failed == 0 and (succeeded + skipped) > 0
        write_envelope(env, {
            "code": 0 if ok else 1,
            "msg": "ok" if ok else f"{failed} 个失败",
            "results": json.dumps(records, ensure_ascii=False),
        })
    else:
        _pause()
        if failed:
            sys.exit(1)


def _record(source: str, target: str, status: str, message: str, outputs: list) -> dict:
    return {"source": source, "target": target, "status": status, "message": message, "outputs": outputs}


def _output_label(outputs: list) -> str:
    if not outputs:
        return ""
    if len(outputs) == 1:
        return outputs[0]
    folder = str(Path(outputs[0]).parent)
    return f"{folder}（{len(outputs)} 个文件）"


def _print_header(files: list, target: str, out_dir) -> None:
    print()
    print(STYLE.cyan + _hr("=") + STYLE.reset)
    print(f"  {STYLE.bold}文档格式转换{STYLE.reset}  {STYLE.dim}(基于开源项目 DocForge · MIT · 作者 binghai-fly){STYLE.reset}")
    print(f"  目标格式 : {STYLE.bold}{target.upper()}{STYLE.reset}")
    print(f"  输出目录 : {out_dir if out_dir else '（源文件所在目录）'}")
    print(f"  待转换   : {len(files)} 个文件")
    print(STYLE.cyan + _hr("=") + STYLE.reset)
    print()
    print(f"  {STYLE.dim}开始转换…{STYLE.reset}")
    print()


def _print_summary(success: int, fail: int, skip: int, total: int, elapsed: float) -> None:
    print()
    print(STYLE.cyan + _hr("-") + STYLE.reset)
    parts = [
        f"{STYLE.green}成功 {success}{STYLE.reset}",
        f"{STYLE.red}失败 {fail}{STYLE.reset}" if fail else f"失败 {fail}",
        f"跳过 {skip}" if skip else f"跳过 {skip}",
    ]
    print(f"  共 {total} 个文件： " + "， ".join(parts) + f"    耗时 {_fmt_duration(elapsed)}")
    print(f"  {_progress_bar(success + fail + skip, total)}")
    if fail == 0 and total > 0:
        print(f"  {STYLE.green}全部完成！{STYLE.reset}")
    else:
        print(f"  {STYLE.yellow}部分文件未转换成功，请查看上方失败原因。{STYLE.reset}")
    print(STYLE.cyan + _hr("-") + STYLE.reset)


def _pause() -> None:
    if PAUSE_SECONDS <= 0 or os.environ.get("DOCFORGE_NO_PAUSE") == "1":
        return
    print(f"\n{STYLE.dim}窗口将在 {PAUSE_SECONDS} 秒后自动关闭…{STYLE.reset}")
    for remaining in range(PAUSE_SECONDS, 0, -1):
        print(f"  {remaining} 秒", end="\r", flush=True)
        time.sleep(1)
    print(" " * 12, end="\r")


if __name__ == "__main__":
    main()
