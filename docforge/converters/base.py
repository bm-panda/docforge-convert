"""转换器注册机制与公共基类。

每个转换器是 ``(源格式, 目标格式) -> 函数`` 的映射。
函数签名约定：``convert(src: Path, dst: Path, options: dict) -> None``。
成功时正常返回；失败时抛出 :class:`ConversionError`。
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, Optional, Tuple

# (源格式, 目标格式) -> 转换函数
REGISTRY: Dict[Tuple[str, str], Callable] = {}

# 转换函数声明的能力元信息：src -> set(dst)
CAPABILITIES: Dict[str, set] = {}


def register(src: str, dst: str):
    """注册一个转换器。可叠加使用以注册多目标。"""

    def decorator(fn: Callable):
        key = (src, dst)
        if key in REGISTRY:
            raise ValueError(f"转换器重复注册: {src} -> {dst}")
        REGISTRY[key] = fn
        CAPABILITIES.setdefault(src, set()).add(dst)
        return fn

    return decorator


class ConversionError(Exception):
    """转换失败时抛出的异常，message 面向用户。"""

    def __init__(self, message: str, *, needs_office: bool = False):
        super().__init__(message)
        self.message = message
        self.needs_office = needs_office  # 是否因缺少 Office 组件而失败


class ConversionOptions(dict):
    """转换选项，行为类似 dict，但提供默认值便捷访问。"""

    def get_opt(self, key: str, default=None):
        return self.get(key, default)


def get_converter(src: str, dst: str) -> Optional[Callable]:
    """根据格式名查找转换器。"""
    return REGISTRY.get((src, dst))


def supported_destinations(src: str) -> set:
    """返回某格式所有可转换的目标格式集合。"""
    return set(CAPABILITIES.get(src, set()))


def is_convertible(src: str, dst: str) -> bool:
    return (src, dst) in REGISTRY
