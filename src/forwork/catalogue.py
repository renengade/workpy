#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件清单生成工具 — catalogue.py
================================

功能
    列出目标文件夹内指定类型的文件（或子文件夹），生成清单并可导出为 Excel。

用法
    python -m forwork catalogue xlsx                  # 列出包目录下所有 xlsx 并导出
    python -m forwork catalogue pdf -d "D:/Desktop"   # 指定目录 + 文件类型
    python -m forwork catalogue folder --no-export    # 仅列出子文件夹，不导出
    python catalogue.py xlsx                            # 也可直接运行本脚本

参数
    file_type            文件类型，如 xlsx / pdf / folder（folder 表示子文件夹）
    -d, --dir            目标文件夹（默认：脚本所在目录）
    -n, --no-export      仅打印到终端，不生成 Excel（默认导出）

交互模式
    不传 file_type 时进入交互模式，依次输入文件夹路径、文件类型、是否导出。
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

import pandas as pd


def get_catalogue(folder: str, file_type: str, export: bool = True) -> pd.DataFrame:
    """返回文件夹内指定类型的文件清单（不含扩展名）。"""
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
    return df


def interactive_mode() -> None:
    """无 file_type 时进入交互模式。"""
    default_folder = os.path.dirname(os.path.abspath(__file__))
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

    df = get_catalogue(folder, file_type, export=export)
    if export:
        timestamp = datetime.now().strftime("%m%d %H%M%S")
        out_path = os.path.join(folder, f"FileName_{timestamp}.xlsx")
        df.to_excel(out_path, index=False)
        print(f"Successfully export file!!! → {out_path}")


def configure_parser(p: argparse.ArgumentParser) -> None:
    """向解析器添加本工具的参数。"""
    p.add_argument("file_type", nargs="?", help="文件类型（如 xlsx / pdf / folder）")
    p.add_argument("-d", "--dir", default=None, help="目标文件夹（默认：脚本所在目录）")
    p.add_argument("-n", "--no-export", action="store_true", help="仅打印，不导出")


def run(args: argparse.Namespace) -> int:
    """执行清单生成，返回退出码。"""
    folder = args.dir or os.path.dirname(os.path.abspath(__file__))
    if not os.path.isdir(folder):
        print(f"❌ 路径不存在: {folder}")
        return 1

    if not args.file_type:
        interactive_mode()
        return 0

    df = get_catalogue(folder, args.file_type)
    if not args.no_export:
        timestamp = datetime.now().strftime("%m%d %H%M%S")
        out_path = os.path.join(folder, f"FileName_{timestamp}.xlsx")
        df.to_excel(out_path, index=False)
        print(f"Successfully export file!!! → {out_path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="catalogue",
        description="生成文件夹内指定类型文件/子文件夹的清单，可导出 Excel",
    )
    configure_parser(parser)
    return run(parser.parse_args(argv))


if __name__ == "__main__":
    raise SystemExit(main())
