#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""forwork 统一命令行入口（子命令调度器）。

【调用方式】
  WSL (Linux)::

      forwork <command> [options]          # 已 uv tool install -e . 后，任意目录可用
      python -m forwork <command> [options] # 在项目目录内直接运行，无需安装
      uv run forwork <command> [options]    # 在项目目录内用项目虚拟环境运行

  Windows::

      forwork <command> [options]          # 已 uv tool install -e . 后，任意目录可用
      forwork.exe <command> [options]      # PowerShell / cmd 中同理

【平台说明】
  - ``catalogue`` / ``merge`` / ``sample`` / ``repics`` 为纯 Python 实现，Linux 与 Windows 均可运行。
  - ``inspics`` / ``itmwp`` 通过 xlwings 调用本机 Microsoft Excel，**只能在 Windows 上运行**；
    在 Linux 上调用会给出明确提示而不是抛异常堆栈。

【路径写法】
  - Windows: ``D:/Work/底稿`` 或 ``D:\\Work\\底稿``
  - WSL 访问 Windows 磁盘: ``/mnt/d/Work/底稿``

各子命令详细参数: ``forwork <command> -h``。
"""

from __future__ import annotations

import argparse
import importlib

from forwork import __version__

# 子命令名 -> (模块路径, 简短说明)
_COMMANDS = {
    "catalogue": ("forwork.catalogue", "生成文件夹内指定类型文件/子文件夹清单"),
    "merge": ("forwork.datamerge", "合并目录下 csv/xlsx 等文件为多种格式"),
    "sample": ("forwork.sampling", "对 xlsx 做可复现随机抽样并导出"),
    "inspics": ("forwork.inspics", "把图片按序号批量插入 Excel（需 Windows + Excel）"),
    "repics": ("forwork.repics", "批量缩放图片尺寸"),
    "itmwp": ("forwork.itmwp", "IT 审计工作底稿批量生成（需 Windows + Excel）"),
}


def build_parser() -> argparse.ArgumentParser:
    """构建顶层解析器并注册全部子命令。

    子命令模块在这里被导入只是为了读取它们的参数定义；
    xlwings（仅 inspics / itmwp 需要）已改为在函数内部延迟导入，
    因此在没有 xlwings 的 Linux 上 ``forwork --help`` 依然可用。
    """
    parser = argparse.ArgumentParser(
        prog="forwork",
        description="Derrick 的工作小工具库",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "各子命令详细参数: forwork <command> -h\n"
            "示例(Windows): forwork itmwp enterprise derr leader 2026.7.27 2026 ELC SD OM CM\n"
            "示例(WSL)    : forwork merge -d /mnt/d/Work/data -t csv -o merged.db"
        ),
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    for name, (module_path, help_text) in _COMMANDS.items():
        module = importlib.import_module(module_path)
        sp = subparsers.add_parser(
            name,
            help=help_text,
            description=module.__doc__,
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )
        module.configure_parser(sp)
        sp.set_defaults(_entry=module.run)
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI 主入口，返回进程退出码（0 成功，非 0 失败）。"""
    parser = build_parser()
    args = parser.parse_args(argv)
    entry = getattr(args, "_entry", None)
    if entry is None:
        parser.print_help()
        return 1
    return entry(args) or 0


if __name__ == "__main__":
    raise SystemExit(main())
