# forwork — Derrick 的工作小工具库

一组面向日常审计 / 办公自动化的小工具，统一通过一个命令行入口调用，也可作为 Python 库导入使用。

所有工具都遵循同一套约定，既可以 `python -m forwork <command>` 统一调用，也可以直接运行单个脚本 `python src/forwork/<tool>.py`。

## 工具清单

| 子命令 | 模块 | 功能 | 主要依赖 |
|--------|------|------|----------|
| `catalogue` | `catalogue.py` | 列出文件夹内指定类型文件 / 子文件夹，导出 Excel 清单 | pandas, openpyxl |
| `merge` | `datamerge.py` | 合并目录下 csv/tsv/txt/xlsx，输出为 sqlite/xlsx/csv/tsv/json | pandas, openpyxl |
| `sample` | `sampling.py` | 纯标准库读取 xlsx，可复现随机抽样并导出 | 纯标准库 |
| `inspics` | `inspics.py` | 把图片按序号批量插入 Excel，生成汇总表 + 超链接 | xlwings, Pillow |
| `repics` | `repics.py` | 批量按比例缩放图片，保存到输出子目录 | Pillow |
| `itmwp` | `itmwp.py` | IT 审计工作底稿批量生成 | xlwings |

## 安装

```bash
# 使用 uv（推荐）
uv sync

# 或使用 pip
pip install .
```

安装后会得到 `forwork` 命令；不安装也可用 `python -m forwork` 运行。

## 快速开始

```bash
# 查看总帮助
python -m forwork --help

# 查看某个子命令的参数
python -m forwork merge --help
```

## 各工具用法

### catalogue — 文件清单

```bash
python -m forwork catalogue xlsx                  # 列出包目录下所有 xlsx 并导出
python -m forwork catalogue pdf -d "D:/Desktop"   # 指定目录 + 文件类型
python -m forwork catalogue folder --no-export    # 仅列出子文件夹，不导出
```

### merge — 数据合并

```bash
# 合并但不写出，仅统计
python -m forwork merge -d ./data -t csv

# 合并并一次导出多种格式
python -m forwork merge -d ./data -t csv -p "日志*" \
    -o out.db -T logs -o out.xlsx -o out.csv
```

### sample — 随机抽样

```bash
# 非交互：指定 sheet、样本量、随机种子、列，并导出
python -m forwork sample 数据.xlsx -s Sheet1 -n 10 -S 42 -c A列,B列 -o -p 结果.xlsx

# 交互：不传 -s/-n/-S/-c 时按提示逐步输入
python -m forwork sample 数据.xlsx
```

### inspics — 图片插入 Excel

```bash
python -m forwork inspics --path D:/photos --picwidth 15 --interval 4
```

> 依赖 Excel（xlwings 调用本机 Excel），目前主要在 Windows 上使用。

### repics — 图片缩放

```bash
python -m forwork repics -d D:/photos -o 缩略图 --wscale 0.5 --hscale 0.5
```

### itmwp — IT 审计底稿生成

```bash
python -m forwork itmwp huada 蒋jj limoumou 2026.7.27 2026 ELC SD OM CM
```

> 依赖 Excel（xlwings），目前主要在 Windows 上使用。

## 作为库调用

```python
from forwork import get_catalogue
from forwork import XlsxReader, random_sample_from_reader, write_sample_xlsx

# 文件清单
df = get_catalogue("D:/Desktop", "xlsx", export=False)

# 纯标准库读取 + 可复现抽样
reader = XlsxReader("数据.xlsx")
rows = random_sample_from_reader(reader, "Sheet1", n=10, seed=42)
```

## 项目结构

```
src/forwork/
├── __init__.py      # 包入口，暴露公共 API 与 __version__
├── __main__.py      # 支持 python -m forwork
├── cli.py           # 统一子命令调度器
├── catalogue.py     # 文件清单
├── datamerge.py     # 数据合并
├── sampling.py      # 随机抽样（纯标准库）
├── inspics.py       # 图片插入 Excel
├── repics.py        # 图片缩放
└── itmwp.py         # IT 审计底稿生成
```

每个工具模块都导出三个约定函数，便于统一调度与单独测试：

- `configure_parser(parser)` —— 向解析器注册参数
- `run(args)` —— 执行主体，返回退出码
- `main(argv=None)` —— 独立运行入口（内部自建解析器）
