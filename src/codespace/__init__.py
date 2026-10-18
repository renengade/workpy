"""codespace —— 文件清单生成工具集。

公共 API:
    get_catalogue(folder, file_type) -> pd.DataFrame

命令行:
    uv run codespace [file_type] [-d DIR] [-n]
"""

from codespace.get_catalogue import get_catalogue, main

__all__ = ["get_catalogue", "main"]
