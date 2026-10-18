"""forwork —— Derrick 的工作小工具库。

命令行（统一子命令）::

    python -m forwork catalogue   生成文件/文件夹清单
    python -m forwork merge       合并 csv/xlsx 等文件
    python -m forwork sample      xlsx 随机抽样
    python -m forwork inspics     图片批量插入 Excel
    python -m forwork repics      批量缩放图片
    python -m forwork mwp         IT 审计工作底稿生成

作为库使用的公共 API::

    from forwork import get_catalogue
    from forwork import XlsxReader, random_sample_from_reader, write_sample_xlsx
    from forwork.mwp import myWorkingPaper          # 需要 xlwings
    from forwork.inspics import insert_images_to_excel  # 需要 xlwings
"""

from __future__ import annotations

__version__ = "0.1.0"

from forwork.catalogue import get_catalogue
from forwork.sampling import (
    XlsxReader,
    random_sample_from_reader,
    write_sample_xlsx,
)

__all__ = [
    "__version__",
    "get_catalogue",
    "XlsxReader",
    "random_sample_from_reader",
    "write_sample_xlsx",
]
