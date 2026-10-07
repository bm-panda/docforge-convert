"""转换器包：导入本模块即完成全部格式的转换注册。"""

from __future__ import annotations

# 导入顺序即注册顺序；各模块通过 @register 填充 base.REGISTRY
from docforge.converters import (  # noqa: F401
    ebook_conv,
    legacy_conv,
    office_com,
    office_conv,
    pdf_conv,
    spreadsheet_conv,
    text_conv,
)

from docforge.converters.base import (  # noqa: F401
    CAPABILITIES,
    REGISTRY,
    ConversionError,
    ConversionOptions,
    get_converter,
    is_convertible,
    supported_destinations,
)
