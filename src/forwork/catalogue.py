#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""catalogue — 文件夹内指定类型文件 / 子文件夹清单，并导出 Excel。

【调用方式】
  WSL (Linux)::

      forwork catalogue xlsx                        # 列出当前目录下的 xlsx 并导出清单
      forwork catalogue pdf -d /mnt/d/Work/底稿      # 指定目录（WSL 访问 Windows 磁盘）
      forwork catalogue folder --no-export           # 仅列出子文件夹，不导出
      python -m forwork catalogue xlsx               # 项目目录内直接运行
      python src/forwork/catalogue.py xlsx           # 直接运行单文件（等价上面一条）

  Windows::

      forwork catalogue xlsx                        # 列出当前目录下的 xlsx 并导出清单
      forwork catalogue pdf -d D:/Work/底稿          # 指定目录
      forwork catalogue folder --no-export           # 仅列出子文件夹，不导出

  库调用::

      from forwork import get_catalogue
      df = get_catalogue("D:/Work/底稿", "xlsx", export=False)

【平台要求】
  纯 Python + pandas，Linux / Windows 均可运行，无需 Excel。

【路径规则】
  - Windows: ``D:/Work/底稿``
  - WSL 访问 Windows 磁盘: ``/mnt/d/Work/底稿``；访问 WSL 自身目录直接用 ``/home/...``

【默认值】
  不传 ``-d`` 时扫描**当前工作目录**（不是脚本所在目录），因此全局安装后
  ``forwork catalogue xlsx`` 会列出你当前所在文件夹的内容。

【输出】
  清单导出为 ``FileName_MMDD HHMMSS.xlsx``（含空格），保存在被扫描的目录内；
  该文件本身也会在下一次扫描时被列出，可用 ``-d`` 指定其它目录避免污染。
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

import pandas as pd


def get_catalogue(folder: str, file_type: str, export: bool = True) -> pd.DataFrame:
    """返回文件夹内指定类型的文件清单（不含扩展名），可选导出为 Excel。

    Args:
        folder:   目标文件夹路径。Windows 形如 ``D:/Work/底稿``；
                  WSL 形如 ``/mnt/d/Work/底稿``。
        file_type: 文件类型（不含点），如 ``xlsx`` / ``pdf`` / ``docx``；
                   传 ``folder`` 则列出子文件夹。
        export:   是否把清单导出为 Excel（``FileName_MMDD HHMMSS.xlsx``，
                   写入 ``folder`` 内）。为 ``False`` 时只返回并打印 DataFrame。

    Returns:
        pandas.DataFrame: 单列 DataFrame，列名即 ``file_type``
        （``folder`` 模式下为 ``folder``），内容为不含扩展名的名称列表。

    Raises:
        FileNotFoundError: ``folder`` 不存在时由 ``os.listdir`` 抛出。

    Example:
        >>> df = get_catalogue("D:/Work/底稿", "xlsx", export=False)
        >>> df.columns.tolist()
        ['xlsx']
    """
    file_type = file_type.lower().lstrip(".")
    entries = sorted(os.listdir(folder))

    if file_type == "folder":
        file_list = [
            item
            for item in entries
            if os.path.isdir(os.path.join(folder, item))
            and not item.startswith(".")
            and item != "__pycache__"
        ]
        col_name = "folder"
    else:
        ext = f".{file_type}"
        file_list = [
            os.path.splitext(x)[0]
            for x in entries
            if os.path.isfile(os.path.join(folder, x))
            and x.lower().endswith(ext)
            and not x.startswith(".")
            and not x.startswith("~$")
        ]
        col_name = file_type

    df = pd.DataFrame(file_list, columns=[col_name])
    print(df)

    if export:
        timestamp = datetime.now().strftime("%m%d %H%M%S")
        out_path = os.path.join(folder, f"FileName_{timestamp}.xlsx")
        df.to_excel(out_path, index=False)
        print(f"Successfully export file!!! → {out_path}")

    return df


def interactive_mode() -> None:
    """交互模式：未提供文件类型时逐项提问。

    提示顺序: 目标文件夹 → 文件类型 → 是否导出 Excel。
    目标文件夹默认当前工作目录。
    """
    default_folder = os.getcwd()
    folder = (
        input(f"请输入文件夹路径（默认: {default_folder}）: ").strip() or default_folder
    )

    if not os.path.isdir(folder):
        print(f"❌ 路径不存在: {folder}")
        sys.exit(1)

    file_type = input("请输入文件类型（如 xlsx/pdf/folder）: ").strip().lower()
    if not file_type:
        print("❌ 文件类型不能为空")
        sys.exit(1)

    answer = input("是否导出清单到 Excel？(y/n，默认 y): ").strip().lower()
    export = answer != "n"

    get_catalogue(folder, file_type, export=export)


def configure_parser(p: argparse.ArgumentParser) -> None:
    """向解析器添加本工具的参数。"""
    p.add_argument("file_type", nargs="?", help="文件类型（如 xlsx / pdf / folder）")
    p.add_argument(
        "-d",
        "--dir",
        default=None,
        help="目标文件夹（默认：当前工作目录）",
    )
    p.add_argument("-n", "--no-export", action="store_true", help="仅打印，不导出")


def run(args: argparse.Namespace) -> int:
    """执行清单生成，返回退出码（0 成功，1 路径不存在）。"""
    folder = args.dir or os.getcwd()
    if not os.path.isdir(folder):
        print(f"❌ 路径不存在: {folder}")
        return 1

    if not args.file_type:
        interactive_mode()
        return 0

    get_catalogue(folder, args.file_type, export=not args.no_export)
    return 0


def main(argv: list[str] | None = None) -> int:
    """独立运行入口（不经过 forwork 统一调度）。"""
    parser = argparse.ArgumentParser(
        prog="catalogue",
        description="生成文件夹内指定类型文件/子文件夹的清单，可导出 Excel",
    )
    configure_parser(parser)
    return run(parser.parse_args(argv))


if __name__ == "__main__":
    raise SystemExit(main())
