#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""inspics — 把文件夹里的图片按序号批量插入 Excel，并生成带超链接的汇总表。

【调用方式】
  Windows（**必须**在 Windows 上运行，需要本机安装 Microsoft Excel）::

      forwork inspics --path D:/Work/照片                        # 处理指定目录
      forwork inspics --path D:/Work/照片 --visible              # 显示 Excel 窗口
      forwork inspics --path D:/Work/照片 --picwidth 15 --interval 4
      python -m forwork inspics --path D:/Work/照片
      python src/forwork/inspics.py --path D:/Work/照片

  WSL (Linux)::

      # 不可用：本工具通过 xlwings 调用本机 Excel，Excel 只装在 Windows 上。
      # 在 WSL 里执行会得到明确提示，而不是一堆异常堆栈。

  库调用::

      from forwork.inspics import insert_images_to_excel
      insert_images_to_excel(Path("D:/Work/照片"), visible=False, picwidth=15, interval=4)

【平台要求】
  需要: Windows + Microsoft Excel + xlwings + Pillow + pandas。
  xlwings 是**延迟导入**的，因此在没有 xlwings 的机器上执行 ``forwork --help``
  仍然正常；只有真正调用本工具时才会提示缺少依赖。

【路径规则】
  Windows: ``D:/Work/照片``（也接受 ``D:\\Work\\照片``）。

【默认值】
  - 不传 ``--path`` 时处理**当前工作目录**下的图片。
  - ``--picwidth`` 图片宽度（厘米，默认 20）；``--interval`` 相邻图片间隔行数（默认 6）。
  - 输出文件 ``outputYYYYMMDD_HHMMSS.xlsx`` 保存在图片所在目录。

【文件命名规则】
  取文件名开头的数字段作为序号（``01-01 测试图.png`` → 序号 ``01-01``，名称 ``测试图``），
  按序号升序排列；无编号的图片会被跳过并提示。
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Sequence

import pandas as pd
from PIL import Image

PIC_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")

# xlwings 的 Picture.width / .height / left / top 单位是「磅(point)」，不是像素。
# 1 英寸 = 72 磅 = 2.54 厘米 → 1 厘米 ≈ 28.3465 磅。
POINTS_PER_CM = 28.3464567

# 图片所在行的行高（磅），与下面 ws1.range("a:a").row_height 保持一致。
ROW_HEIGHT_PT = 15


# ====================================================================
#  命令行参数解析
# ====================================================================


def configure_parser(parser: argparse.ArgumentParser) -> None:
    """向解析器添加本工具的参数。"""
    parser.add_argument(
        "--path",
        "-p",
        default=None,
        help="图片所在文件夹（默认：当前工作目录）",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="显示 Excel 窗口（默认后台运行）",
    )
    parser.add_argument(
        "--picwidth",
        type=float,
        default=20.0,
        help="图片宽度，单位 cm（默认 20）",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=6,
        help="相邻图片之间的间隔行数（默认 6）",
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="inspics",
        description="将文件夹中的图片按序号批量插入 Excel，汇总表单独放在一个 sheet",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例(Windows): forwork inspics --path D:/photos --picwidth 15 --interval 4\n"
            "注意: 需要 Windows + 本机安装 Microsoft Excel"
        ),
    )
    configure_parser(parser)
    return parser


# ====================================================================
#  工具函数
# ====================================================================


def _split_seq(filename: str) -> tuple[tuple[int, ...], str, str] | None:
    """从文件名中拆分编号和名称文本。

    编号 = 文件名开头的数字段（如 "01-01"），
    名称 = 编号之后、扩展名之前的文本部分（如 "测试图"）。
    无编号时返回 None。

    例:
      "01-01 测试图.png" -> ((1, 1), "01-01", "测试图")
      "02-系统截图.jpg"  -> ((2,), "02", "系统截图")
      "现场照片.png"     -> None
    """
    m = re.match(r"\d+(?:-\d+)*", filename)
    if not m:
        return None
    seq_str = m.group()
    sort_key = tuple(int(x) for x in seq_str.split("-"))
    rest = Path(filename).stem[m.end() :]
    name = rest.lstrip("- _.").strip()
    return sort_key, seq_str, name


def _collect_images(folder: Path) -> list[tuple[tuple[int, ...], str, str, Path]]:
    """收集文件夹中带编号的图片文件，按编号排序。

    Returns:
        按编号升序排列的 [(排序键, 编号串, 名称文本, 文件路径), ...]，
        无编号的图片跳过并打印警告。
    """
    images: list[tuple[tuple[int, ...], str, str, Path]] = []
    for f in folder.iterdir():
        if not (f.is_file() and f.suffix.lower() in PIC_EXTENSIONS):
            continue
        parsed = _split_seq(f.name)
        if parsed is None:
            print(f"跳过（文件名无编号）: {f.name}")
            continue
        images.append((parsed[0], parsed[1], parsed[2], f))

    images.sort(key=lambda t: (t[0], t[3].name.lower()))
    return images


def _import_xlwings():
    """延迟导入 xlwings，并在不可用时给出可读的错误提示。

    Returns:
        module: xlwings 模块。

    Raises:
        SystemExit: 未安装 xlwings 或当前系统不支持 Excel 自动化时退出。
    """
    from forwork._excel import require_xlwings

    return require_xlwings()


# ====================================================================
#  主流程
# ====================================================================


def insert_images_to_excel(
    folder: Path,
    visible: bool = False,
    picwidth: float = 20.0,
    interval: int = 6,
) -> None:
    """将 folder 中带序号的图片按序插入 Excel。

    Args:
        folder:   图片所在文件夹（输出文件也保存在这里）。Windows 形如
                  ``Path("D:/Work/照片")``。
        visible:  是否显示 Excel 窗口（默认后台运行）。
        picwidth: 图片宽度（**厘米**），高度按原图比例缩放。
        interval: 相邻图片之间的间隔行数。

    输出:
        在 ``folder`` 内生成 ``outputYYYYMMDD_HHMMSS.xlsx``，
        内含「图片」sheet 与「汇总表」sheet（汇总表的序号列带内部超链接）。

    注意:
        需要 Windows + Microsoft Excel。执行过程中会启动一个隐藏的 Excel 进程，
        无论成功或异常都会在结束时关闭。
    """
    xw = _import_xlwings()

    images = _collect_images(folder)
    if not images:
        print("未找到带序号的图片文件，程序退出")
        return

    app = None
    try:
        app = xw.App(visible=visible, add_book=False)
        wb = app.books.add()
        ws1 = wb.sheets.active  # 「图片」sheet
        ws1.range("a:a").row_height = ROW_HEIGHT_PT  # 行高（磅）
        ws2 = wb.sheets.add("汇总表", after=ws1)  # 汇总表单独一个 sheet，放在图片之后

        # === 插入图片并记录每张图的实际位置 ===
        row = 6  # 图片起始行
        results: list[tuple[str, str, str]] = []  # (位置, 序号, 图片名称)

        for _, seq_str, pic_name, image_path in images:
            # 图片上方一行标注图片名称
            ws1.range(f"a{row - 1}").value = pic_name

            with Image.open(image_path) as img:
                w, h = img.width, img.height

            # 宽度按 picwidth(厘米) 换算成磅，高度等比缩放
            width_pt = POINTS_PER_CM * picwidth
            ws1.pictures.add(
                str(image_path),
                name=image_path.stem,
                update=True,
                left=0,
                top=ws1.range(f"a{row}").top,
                width=width_pt,
                height=width_pt * h / w,
            )

            # 图片占用的行数 = 图片高度(磅) / 行高(磅)
            pic_height_pt = ws1.pictures[-1].height
            rows_used = math.ceil(pic_height_pt / ROW_HEIGHT_PT)

            results.append((f"A{row}", seq_str, pic_name))
            row += rows_used + interval

        # === 在独立的「汇总表」sheet 中写入汇总表 ===
        ws2.range("a1").value = ["位置", "序号", "图片名称"]
        ws2.range("a1:c1").font.bold = True
        # 序号列按文本格式写入，避免 "01-01" 被 Excel 识别为日期
        ws2.range(f"b2:b{1 + len(results)}").number_format = "@"
        for i, (addr, seq, name) in enumerate(results):
            row = 2 + i
            ws2.range(f"a{row}").value = [addr, seq, name]
            # 序号单元格添加内部超链接：点击跳转到图片所在位置
            cell = ws2.range(f"b{row}")
            cell.api.Hyperlinks.Add(
                Anchor=cell.api,
                Address="",
                SubAddress=f"'{ws1.name}'!{addr}",
            )
        ws2.autofit()

        # === 保存输出（输出文件与图片同目录） ===
        date = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = folder / f"output{date}.xlsx"
        wb.save(str(out_path))

        summary = pd.DataFrame(results, columns=["位置", "序号", "图片名称"])
        print(summary)
        print(f"export successful! -> {out_path}")
        wb.close()
    finally:
        # 无论成功或异常，都必须关闭 Excel 进程，否则会残留 EXCEL.EXE
        if app is not None:
            try:
                app.quit()
            except Exception as e:  # noqa: BLE001
                print(f"[警告] 关闭 Excel 时出错: {e}")


# ====================================================================
#  程序入口
# ====================================================================


def run(args: argparse.Namespace) -> int:
    """执行图片插入，返回退出码（0 成功，1 目录不存在）。"""
    folder = Path(args.path).resolve() if args.path else Path.cwd()
    if not folder.is_dir():
        print(f"错误：目录不存在 - {folder}")
        return 1

    insert_images_to_excel(
        folder=folder,
        visible=args.visible,
        picwidth=args.picwidth,
        interval=args.interval,
    )
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """独立运行入口（不经过 forwork 统一调度）。"""
    return run(_build_parser().parse_args(argv))


if __name__ == "__main__":
    sys.exit(main())
