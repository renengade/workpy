#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel 自动化相关的公共辅助（延迟导入 xlwings + 平台检查）。

只有 ``inspics``（图片插入 Excel）与 ``itmwp``（IT 审计底稿）需要本机 Excel。
把 xlwings 的导入集中放在这里，可以做到：

- 在**没有 xlwings 的机器**（例如纯 Linux 环境）上执行 ``forwork --help``
  依然完全正常，不会在导入阶段就崩掉；
- 在**不支持 Excel 自动化**的系统上给出可读的中文提示，而不是抛英文异常堆栈。
"""

from __future__ import annotations

import sys


def require_xlwings():
    """导入 xlwings 并确认当前系统支持 Excel 自动化。

    Returns:
        module: xlwings 模块，可直接使用 ``xw.App(...)``。

    Raises:
        SystemExit: 未安装 xlwings，或当前系统不支持 Excel 自动化时退出，
            并打印可读的中文提示。
    """
    if sys.platform.startswith("linux"):
        raise SystemExit(
            "本工具需要 Windows + 本机安装的 Microsoft Excel。\n"
            "当前系统是 Linux —— xlwings 无法在此调用 Excel。\n"
            "请在 Windows 上运行（WSL 里可访问 /mnt/c/... 与 /mnt/d/...，"
            "但 Excel 自动化必须回到 Windows 侧执行）。"
        )

    try:
        import xlwings as xw
    except ImportError as e:
        raise SystemExit(
            "缺少依赖 xlwings。请在 Windows 上执行：pip install xlwings"
            f"（当前错误: {e}）"
        ) from e

    if getattr(xw, "apps", None) is None:
        raise SystemExit(
            "xlwings 未能连接到本机 Excel。请确认已安装 Microsoft Excel"
            "（本工具通过 COM 调用 Excel，需要本机安装）。"
        )
    return xw
