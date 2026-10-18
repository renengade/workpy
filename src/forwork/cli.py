#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""forwork 统一命令行入口。

用法::

    python -m forwork <command> [options]
    forwork <command> [options]          # 安装后

子命令::

    catalogue   生成文件夹内指定类型文件/子文件夹清单
    merge       合并目录下 csv/xlsx 等文件为多种格式
    sample      对 xlsx 做可复现随机抽样并导出
    inspics     把图片按序号批量插入 Excel
    repics      批量缩放图片尺寸
    itmwp       IT 审计工作底稿批量生成

各子命令详细参数：``python -m forwork <command> -h``。
"""

from __future__ import annotations

import argparse

from forwork import __version__
from forwork import catalogue, datamerge, inspics, itmwp, repics, sampling

# 子命令名 -> (模块, 简短说明)
_COMMANDS = {
    "catalogue": (catalogue, "生成文件夹内指定类型文件/子文件夹清单"),
    "merge": (datamerge, "合并目录下 csv/xlsx 等文件为多种格式"),
    "sample": (sampling, "对 xlsx 做可复现随机抽样并导出"),
    "inspics": (inspics, "把图片按序号批量插入 Excel"),
    "repics": (repics, "批量缩放图片尺寸"),
    "itmwp": (itmwp, "IT 审计工作底稿批量生成"),
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="forwork",
        description="Derrick 的工作小工具库",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="各子命令详细参数: python -m forwork <command> -h",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )
    subparsers = parser.add_subparsers(dest="command", metavar="<command>")
    for name, (module, help_text) in _COMMANDS.items():
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
    parser = build_parser()
    args = parser.parse_args(argv)
    entry = getattr(args, "_entry", None)
    if entry is None:
        parser.print_help()
        return 1
    return entry(args) or 0


if __name__ == "__main__":
    raise SystemExit(main())
