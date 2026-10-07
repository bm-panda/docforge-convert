"""DocForge —— 万档转换器（全能文档格式转换软件）。

DocForge 是一个面向个人与企业用户的文档格式转换工具，
支持文档、表格、演示、电子书、数据等类别多种格式互转。

主要入口：
  - 命令行：  docforge convert <文件> -t <格式>
"""

__version__ = "1.0.1"
APP_NAME = "DocForge 万档转换器"
APP_TAGLINE = "全能文档格式转换软件"

from docforge.core import engine  # noqa: E402,F401
from docforge.core.engine import DocForgeEngine  # noqa: E402,F401
