#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""repics — 批量按比例缩放图片，保存到输出子目录。

【调用方式】
  WSL (Linux)::

      forwork repics -d /mnt/d/Work/照片 -o 缩略图 --wscale 0.5 --hscale 0.5
      forwork repics                       # 处理当前目录下的图片，默认缩放 50%
      python -m forwork repics -d ./照片
      python src/forwork/repics.py -d ./照片

  Windows::

      forwork repics -d D:/Work/照片 -o 缩略图 --wscale 0.5 --hscale 0.5
      forwork repics --wscale 0.3 --hscale 0.3

  库调用::

      from forwork.repics import reduce_size
      reduce_size(Path("D:/Work/照片"), Path("D:/Work/照片/缩略图"), 0.5, 0.5)

【平台要求】
  纯 Python + Pillow，Linux / Windows 均可运行，无需 Excel。

【路径规则】
  - Windows: ``D:/Work/照片``
  - WSL 访问 Windows 磁盘: ``/mnt/d/Work/照片``

【默认值】
  - 不传 ``-d`` 时处理**当前工作目录**下的图片。
  - ``-o`` 输出子目录名（默认 ``调整尺寸``），位于图片目录内部，**不覆盖原图**。
  - ``--wscale`` / ``--hscale`` 宽高缩放比例（默认 0.5）。

【支持的格式】
  .jpg / .jpeg / .png / .bmp / .gif / .webp / .tif / .tiff。
  缩放前会自动按 EXIF 方向摆正（手机竖拍照片不会躺倒）。
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from PIL import Image, ImageOps

PIC_EXTENSIONS = (
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".gif",
    ".webp",
    ".tif",
    ".tiff",
)


def reduce_size(
    src_dir: Path,
    out_dir: Path,
    w_scale: float = 0.5,
    h_scale: float = 0.5,
) -> int:
    """批量缩放 src_dir 中的图片到 out_dir，返回成功处理的图片数量。

    Args:
        src_dir:  源图片目录（只处理该目录下的文件，不递归子目录）。
        out_dir:  输出目录（自动创建）。
        w_scale:  宽度缩放比例（0.5 = 缩小到一半）。
        h_scale:  高度缩放比例。

    Returns:
        int: 成功写出的图片数量。

    注意:
        缩放前会按 EXIF Orientation 摆正图片；单张图片失败会被跳过并打印提示，
        不影响其余图片。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    done = 0
    for pic in sorted(src_dir.iterdir()):
        if not (pic.is_file() and pic.suffix.lower() in PIC_EXTENSIONS):
            continue
        try:
            with Image.open(pic) as im:
                # 按 EXIF 方向摆正（手机竖拍照片），避免缩放后方向错乱
                fixed = ImageOps.exif_transpose(im)
                try:
                    w, h = fixed.size
                    img = fixed.resize(
                        (max(1, int(w * w_scale)), max(1, int(h * h_scale))),
                        Image.Resampling.LANCZOS,
                    )
                    try:
                        img.save(out_dir / pic.name)
                    finally:
                        img.close()
                finally:
                    fixed.close()
            done += 1
        except Exception as e:  # noqa: BLE001
            print(f"[跳过] 处理 {pic.name} 失败: {e}")
    print(f"共缩放 {done} 张图片 -> {out_dir}")
    return done


def configure_parser(p: argparse.ArgumentParser) -> None:
    """向解析器添加本工具的参数。"""
    p.add_argument("-d", "--dir", default=None, help="图片目录（默认：当前工作目录）")
    p.add_argument(
        "-o", "--out", default="调整尺寸", help="输出子目录名（默认：调整尺寸）"
    )
    p.add_argument("--wscale", type=float, default=0.5, help="宽度缩放比例（默认 0.5）")
    p.add_argument("--hscale", type=float, default=0.5, help="高度缩放比例（默认 0.5）")


def run(args: argparse.Namespace) -> int:
    """执行图片缩放，返回退出码（0 成功，1 目录不存在或输出目录非法）。"""
    src_dir = Path(args.dir).resolve() if args.dir else Path.cwd()
    if not src_dir.is_dir():
        print(f"❌ 目录不存在: {src_dir}")
        return 1
    out_dir = src_dir / args.out
    if out_dir.resolve() == src_dir.resolve():
        print("❌ 输出子目录不能与图片目录相同")
        return 1
    reduce_size(src_dir, out_dir, w_scale=args.wscale, h_scale=args.hscale)
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """独立运行入口（不经过 forwork 统一调度）。"""
    parser = argparse.ArgumentParser(
        prog="repics",
        description="批量缩放图片尺寸，结果保存到输出子目录",
    )
    configure_parser(parser)
    return run(parser.parse_args(argv))


if __name__ == "__main__":
    raise SystemExit(main())
