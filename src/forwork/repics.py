#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片批量缩放工具 — repics.py
============================

功能
    将目标文件夹中的图片（.jpg/.jpeg/.png/.bmp/.gif）按比例缩放，
    结果保存到输出子目录（默认「调整尺寸」），不覆盖原图。

用法
    python -m forwork repics                           # 处理脚本所在目录，默认缩放 50%
    python -m forwork repics -d D:/photos -o 缩略图     # 指定目录与输出子目录
    python -m forwork repics --wscale 0.3 --hscale 0.3  # 自定义缩放比例
    python repics.py                                      # 直接运行本脚本

依赖
    pip install Pillow
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from PIL import Image

PIC_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif")


def reduce_size(
    src_dir: Path,
    out_dir: Path,
    w_scale: float = 0.5,
    h_scale: float = 0.5,
) -> int:
    """批量缩放 src_dir 中的图片到 out_dir，返回成功处理的图片数量。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    done = 0
    for pic in sorted(src_dir.iterdir()):
        if not (pic.is_file() and pic.suffix.lower() in PIC_EXTENSIONS):
            continue
        try:
            with Image.open(pic) as im:
                w, h = im.size
                img = im.resize(
                    (max(1, int(w * w_scale)), max(1, int(h * h_scale))),
                    Image.Resampling.LANCZOS,
                )
                img.save(out_dir / pic.name)
            done += 1
        except Exception as e:  # noqa: BLE001
            print(f"[跳过] 处理 {pic.name} 失败: {e}")
    print(f"共缩放 {done} 张图片 -> {out_dir}")
    return done


def configure_parser(p: argparse.ArgumentParser) -> None:
    """向解析器添加本工具的参数。"""
    p.add_argument("-d", "--dir", default=None, help="图片目录（默认：脚本所在目录）")
    p.add_argument(
        "-o", "--out", default="调整尺寸", help="输出子目录名（默认：调整尺寸）"
    )
    p.add_argument("--wscale", type=float, default=0.5, help="宽度缩放比例（默认 0.5）")
    p.add_argument("--hscale", type=float, default=0.5, help="高度缩放比例（默认 0.5）")


def run(args: argparse.Namespace) -> int:
    """执行图片缩放，返回退出码。"""
    src_dir = Path(args.dir).resolve() if args.dir else Path(__file__).resolve().parent
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
    parser = argparse.ArgumentParser(
        prog="repics",
        description="批量缩放图片尺寸，结果保存到输出子目录",
    )
    configure_parser(parser)
    return run(parser.parse_args(argv))


if __name__ == "__main__":
    raise SystemExit(main())
