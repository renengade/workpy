"""forwork —— Derrick 的工作小工具库（文件清单 / 数据合并 / 抽样 / 图片 / 底稿）。

【调用方式】
  WSL (Linux)::

      forwork <command> [options]            # 全局命令（uv tool install -e . 后任意目录可用）
      python -m forwork <command> [options] # 项目目录内直接运行
      uv run forwork <command> [options]    # 项目目录内用项目虚拟环境运行

  Windows::

      forwork <command> [options]            # 全局命令（uv tool install -e . 后任意目录可用）
      forwork.exe <command> [options]

【子命令与平台】
  catalogue  文件/子文件夹清单导出          纯 Python     Linux / Windows
  merge      csv/xlsx 合并为多种格式        纯 Python     Linux / Windows
  sample     xlsx 可复现随机抽样            纯标准库     Linux / Windows
  repics     图片批量缩放                  Pillow        Linux / Windows
  inspics    图片批量插入 Excel            xlwings       仅 Windows（需本机 Excel）
  itmwp      IT 审计底稿批量生成           xlwings       仅 Windows（需本机 Excel）

【作为库使用】
  公共 API::

      from forwork import get_catalogue
      from forwork import XlsxReader, random_sample_from_reader, write_sample_xlsx

  子模块 API（按需导入）::

      from forwork.datamerge import find_files, load_files, write_all
      from forwork.repics import reduce_size
      from forwork.inspics import insert_images_to_excel   # 需要 xlwings
      from forwork.itmwp import myWorkingPaper             # 需要 xlwings
"""

from __future__ import annotations

__version__ = "0.2.0"

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
