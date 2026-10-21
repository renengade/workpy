# forwork — Derrick 的工作小工具库

一组面向日常审计 / 办公自动化的小工具：文件清单、数据合并、可复现抽样、图片处理、IT 审计底稿生成。

所有工具统一通过一个命令行入口 `forwork` 调用，也可以作为 Python 库导入使用。

- **开发在 WSL（Linux）**，**使用在 Windows** —— 两端各装一份 editable 安装，靠 GitHub 同步代码。
- 纯 Python 工具（`catalogue` / `merge` / `sample` / `repics`）**Linux 与 Windows 都能跑**。
- 需要操作 Excel 的工具（`inspics` / `itmwp`）**只能在 Windows 上跑**（xlwings 通过 COM 调用本机 Excel）。

---

## 目录

- [工具清单](#工具清单)
- [环境要求](#环境要求)
- [安装](#安装)
  - [WSL（开发机）](#wsl开发机)
  - [Windows（使用机）](#windows使用机)
  - [两端同步](#两端同步)
  - [常见安装问题](#常见安装问题)
- [快速开始](#快速开始)
- [各工具详解](#各工具详解)
- [作为库使用](#作为库使用)
- [常见问题](#常见问题)
- [项目结构](#项目结构)
- [开发指南](#开发指南)

---

## 工具清单

| 子命令 | 模块 | 功能 | 依赖 | 可运行的系统 |
|--------|------|------|------|--------------|
| `catalogue` | `catalogue.py` | 列出文件夹内指定类型文件 / 子文件夹，导出 Excel 清单 | pandas, openpyxl | Linux / Windows |
| `merge` | `datamerge.py` | 合并目录下 csv/tsv/txt/xlsx，输出为 sqlite/xlsx/csv/tsv/json | pandas, openpyxl | Linux / Windows |
| `sample` | `sampling.py` | 纯标准库读取 xlsx，可复现随机抽样并导出 | **纯标准库** | Linux / Windows |
| `repics` | `repics.py` | 批量按比例缩放图片，保存到输出子目录 | Pillow | Linux / Windows |
| `inspics` | `inspics.py` | 把图片按序号批量插入 Excel，生成汇总表 + 超链接 | xlwings, Pillow, pandas | **仅 Windows**（需本机 Excel） |
| `itmwp` | `itmwp.py` | IT 审计工作底稿批量生成 | xlwings | **仅 Windows**（需本机 Excel） |

> 统一约定：每个模块都导出 `configure_parser(parser)` / `run(args)` / `main(argv=None)` 三个函数，
> 因此既能被 `forwork` 统一调度，也能单独测试。

---

## 环境要求

| 项目 | 要求 | 说明 |
|------|------|------|
| Python | **>= 3.11** | 开发环境用 3.15（见 `.python-version`）；Windows 端 uv 会自动准备 |
| uv | 推荐 | 用来管理虚拟环境与全局安装（[官网](https://docs.astral.sh/uv/)） |
| Microsoft Excel | 仅 `inspics` / `itmwp` 需要 | 必须是**本机安装**的 Excel，xlwings 通过 COM 调用 |
| Git | 两端同步需要 | 远端仓库：<https://github.com/renengade/workpy> |

依赖清单（见 `pyproject.toml`）：`openpyxl`、`pandas`、`pillow`、`xlwings`。

---

## 安装

### WSL（开发机）

WSL 是你的**主战场**：改代码、跑测试、提交都在这里做。

```bash
# 1) 拉取代码
cd ~/projects
git clone https://github.com/renengade/workpy.git forwork
cd forwork

# 2) 同步依赖 + 创建项目虚拟环境 .venv（内含 forwork 的可编辑安装）
uv sync

# 3) 把 CLI 注册成全局命令（任意目录可用）
uv tool install -e /home/derrick/projects/forwork

# 4) 把 ~/.local/bin 加入 PATH
uv tool update-shell          # 自动改 ~/.bashrc
source ~/.bashrc             # 立即生效（或重开终端）
```

验证：

```bash
which forwork     # → /home/derrick/.local/bin/forwork
forwork --version # → forwork 0.2.0
forwork --help
```

> `uv tool install -e` 与 `uv sync` 是**互补**的，不是二选一：
> - `uv sync` 管**开发**：`.venv/` 里装了全部依赖 + 项目本身（可编辑），供 ruff / IDE / 测试使用。
> - `uv tool install -e` 管**全局命令**：在 `~/.local/share/uv/tools/forwork/` 建一份隔离环境，
>   并把 `~/.local/bin/forwork` 放到 PATH，任意目录直接敲 `forwork`。
>
> 因为是**可编辑安装**，改完 `.py` 立即生效，不用重新安装。

在 WSL 里访问 Windows 磁盘：`/mnt/d/Work/...`（D 盘）、`/mnt/c/...`（C 盘）。

### Windows（使用机）

Windows 侧是**部署副本**：只在这里运行，不要改代码（改了不会同步回 WSL 的仓库）。

```powershell
# 1) 装依赖（winget 源里都有）
winget install --id Git.Git -e
winget install --id astral-sh.uv -e

# 2) 克隆代码（公开仓库，无需认证）
git clone https://github.com/renengade/workpy.git D:\Code\forwork
cd D:\Code\forwork

# 3) 注册全局命令（uv 会自动准备 Python 解释器并安装依赖）
uv tool install -e .

# 4) 把 %USERPROFILE%\.local\bin 加入「用户级 PATH」（重开终端生效）
[Environment]::SetEnvironmentVariable('Path',
  [Environment]::GetEnvironmentVariable('Path','User') + ';C:\Users\Derrick\.local\bin','User')
```

验证：

```powershell
cd C:\Users\Derrick
forwork --version   # → forwork 0.2.0
forwork --help
```

> `uv tool install -e .` 在 Windows 上会：
> 1. 找到或自动下载一个满足 `requires-python >= 3.11` 的解释器；
> 2. 在 `%APPDATA%\uv\tools\forwork\` 建独立环境并装好依赖；
> 3. 把 `%USERPROFILE%\.local\bin\forwork.exe` 放到 PATH。
>
> 同样因为是**可编辑安装**，`git pull` 之后新代码立刻生效，不需要重新安装。

### 两端同步

```
WSL（改代码）                          Windows（用工具）
──────────────────────────            ──────────────────────────
git add -A                              git pull
git commit -m "..."        ──push──▶    （editable 安装 → 立刻生效）
```

```bash
# WSL 侧
git add -A && git commit -m "fix: merge 错误提示" && git push
```
```powershell
# Windows 侧
cd D:\Code\forwork; git pull
```

日常维护命令：

```bash
uv tool list --show-paths        # 看装到哪了
uv tool upgrade forwork           # 依赖变了 → 重新解析安装
uv tool install -e . --reinstall  # 同上（强制重来）
uv tool uninstall forwork         # 卸载
```

> ⚠️ WSL 直连 GitHub 偶尔 TLS 握手卡死，`git push` 失败时**重试**即可；Windows 侧网络通畅，不受影响。

### 常见安装问题

| 现象 | 原因 / 解决 |
|------|------------|
| `forwork: command not found` | `~/.local/bin` 不在 PATH。执行 `uv tool update-shell` 后重开终端 |
| Windows 提示 `'forwork' 不是内部或外部命令` | `%USERPROFILE%\.local\bin` 没加进用户 PATH，重开终端 |
| `uv tool install` 卡在下载 | PyPI 走官方源慢。WSL 侧已在 `~/.config/uv/uv.toml` 配了清华镜像；Windows 侧可加 `UV_DEFAULT_INDEX` |
| `requires-python >= 3.11` 不满足 | `uv tool install -e . --python 3.13` 指定解释器，或让 uv 自动下载 |
| 改了代码 Windows 没生效 | 确认是 `git pull` 了；`uv tool install -e .` 是可编辑安装，不需要重装 |

---

## 快速开始

三种等价调用方式，按场景选：

```bash
# ① 全局命令（推荐，任意目录）
forwork <command> [options]

# ② 项目内直接运行（开发时用，不依赖全局安装）
cd ~/projects/forwork
python -m forwork <command> [options]     # 用 .venv 里的 python
uv run forwork <command> [options]

# ③ 单个脚本直接运行
python src/forwork/<tool>.py [options]
```

查看帮助：

```bash
forwork --help              # 总帮助
forwork merge --help        # 子命令参数
```

---

## 各工具详解

### catalogue — 文件清单

列出文件夹内指定类型的文件（或子文件夹），并导出为 Excel 清单。

```bash
# 列出当前目录下的 xlsx 并导出清单
forwork catalogue xlsx

# 指定目录（Windows / WSL 路径写法不同）
forwork catalogue pdf -d "D:/Work/底稿"        # Windows
forwork catalogue pdf -d /mnt/d/Work/底稿      # WSL

# 只列出子文件夹，不导出
forwork catalogue folder --no-export

# 不传文件类型 → 进入交互模式（逐项提问）
forwork catalogue
```

| 参数 | 说明 |
|------|------|
| `file_type` | 文件类型（`xlsx` / `pdf` / `folder`…），留空进入交互模式 |
| `-d, --dir` | 目标文件夹，**默认当前工作目录** |
| `-n, --no-export` | 只打印，不导出 |

**输出**：`FileName_MMDD HHMMSS.xlsx`，保存在被扫描的目录内。
注意该文件本身也会在下一次扫描时被列出来，建议用 `-d` 指向专门的资料目录。

### merge — 数据合并

把一个目录下的多个 csv / tsv / txt / xlsx 合并成一张表，一次导出多种格式。

```bash
# 只合并统计，不写文件
forwork merge -d ./data -t csv

# 按文件名过滤 + 指定列 + 一次导出三种格式
forwork merge -d ./data -t csv -p "日志*" -c 时间,用户,内容 \
    -o out.db -T logs -o out.xlsx -o out.csv

# 递归子目录，并记录来源文件
forwork merge -d ./data -t xlsx --recursive --add-source

# 只看会处理哪些文件
forwork merge -d ./data -t csv --dry-run
```

| 参数 | 说明 |
|------|------|
| `-d, --dir` | 扫描目录，**默认当前工作目录** |
| `-t, --type` | 输入类型：`csv` / `tsv` / `txt` / `xlsx`（**不支持 xls**） |
| `-p, --pattern` | 文件名通配符，如 `"日志*"` |
| `-c, --columns` | 只合并这些列，逗号分隔；不传则取所有列的并集 |
| `-o, --output` | 输出路径，可多次指定；格式按扩展名推断 |
| `-f, --format` | 显式指定输出格式（只在单个 `-o` 时有效） |
| `-T, --table` | SQLite 表名（默认取输出文件名） |
| `--if-exists` | SQLite 表已存在时：`replace`（默认）/ `append` / `fail` |
| `--recursive` | 递归扫描子目录 |
| `--add-source` | 增加 `_source_file` 列记录来源文件名 |
| `--encoding` / `--out-encoding` | 输入 / 输出编码（输出默认 `utf-8-sig`，Excel 友好） |
| `--dry-run` | 只列出将要处理的文件 |

**输出格式推断**：`.db/.sqlite/.sqlite3` → SQLite；`.xlsx` → Excel；`.csv` → CSV；`.tsv/.txt` → TSV；`.json` → JSON。

### sample — 随机抽样

**纯标准库**实现（`zipfile` + `xml.etree`），不依赖 pandas / openpyxl，任何装了 Python 的机器都能跑。

```bash
# 非交互：指定 sheet、样本量、随机种子、列，并导出
forwork sample 数据.xlsx -s Sheet1 -n 10 -S 42 -c A列,B列 -o -p 结果.xlsx

# 交互：不传 -s/-n/-S/-c 时按提示逐步输入
forwork sample 数据.xlsx
```

| 参数 | 说明 |
|------|------|
| `file` | `.xlsx` 文件路径（**不支持 .xls**） |
| `-s, --sheet` | sheet 名，留空则交互选择 |
| `-n, --num` | 样本量，留空则交互输入 |
| `-S, --seed` | 随机种子，**同种子结果完全可复现** |
| `-c, --cols` | 只返回这些列（逗号分隔） |
| `-o, --output` | 写出抽样结果为 xlsx |
| `-p, --path` | 输出路径；不传则写入**当前工作目录** |

**输出**：`<原文件名>_<sheet>_sample.xlsx`。
程序先打印每个 sheet 的结构（数据行数、列名、推断类型），再抽样。

### repics — 图片缩放

批量按比例缩放图片，输出到**子目录**，不覆盖原图。

```bash
forwork repics -d D:/photos -o 缩略图 --wscale 0.5 --hscale 0.5
forwork repics                       # 当前目录，默认缩放 50%
```

| 参数 | 说明 |
|------|------|
| `-d, --dir` | 图片目录，**默认当前工作目录** |
| `-o, --out` | 输出子目录名（默认 `调整尺寸`），位于图片目录内 |
| `--wscale` / `--hscale` | 宽 / 高缩放比例（默认 0.5） |

支持 `.jpg/.jpeg/.png/.bmp/.gif/.webp/.tif/.tiff`；缩放前会按 **EXIF 方向摆正**，手机竖拍照片不会躺倒。

### inspics — 图片插入 Excel（仅 Windows）

把带编号的图片按序插入 Excel，生成「图片」sheet + 带超链接的「汇总表」sheet。

```bash
forwork inspics --path D:/photos
forwork inspics --path D:/photos --visible --picwidth 15 --interval 4
```

| 参数 | 说明 |
|------|------|
| `-p, --path` | 图片目录，**默认当前工作目录** |
| `--visible` | 显示 Excel 窗口（默认后台运行） |
| `--picwidth` | 图片宽度（**厘米**，默认 20），高度按原图比例缩放 |
| `--interval` | 相邻图片间隔行数（默认 6） |

**命名规则**：取文件名开头的数字段作序号（`01-01 测试图.png` → 序号 `01-01`、名称 `测试图`），按序号排序；无编号的图片跳过并提示。

**输出**：`outputYYYYMMDD_HHMMSS.xlsx`，保存在图片目录内。

### itmwp — IT 审计底稿生成（仅 Windows）

基于模板批量生成 IT 审计工作底稿。

```bash
forwork itmwp enterprise derr leader 2026.7.27 2026 ELC SD OM CM
forwork itmwp -w D:/Work/2026底稿 enterprise derr leader 2026.7.27 2026 ELC
```

| 位置参数 | 说明 |
|----------|------|
| `client` | 客户简称（替换模板文件名中的 `XXXX`） |
| `name` | 编制人 |
| `reviewer` | 复核人 |
| `prepare` | 编制日期，`YYYY.M.D`（如 `2026.7.27`）；**复核日期自动 = 编制日期 + 1 个月** |
| `audity` | 审计年度（替换模板文件名中的「期间」） |
| `TYPE...` | 底稿类型：`ELC` / `SD` / `CM` / `OM` / `SA`，可多选 |
| `-w, --workdir` | 工作目录，**默认当前工作目录** |

**工作目录**里必须有两个输入：

```
工作目录/
├── FileName.xlsx     支持文档清单（必须有 'folder' 列）
└── ITaudit_v1/       模板底稿目录
```

生成结果也写回该目录：

```
工作目录/{客户简称}_{审计年度}/
├── ITELC_企业层面控制_{客户}_{年度}/ITELC_支持性文档_{客户}_{年度}/
│   ├── ELC01-01 战略规划与执行情况/
│   └── ...
└── ITGC_一般控制测试_{客户}_{年度}/ITGC_支持性文档_{客户}_{年度}/
    ├── SD01-01 可行性分析与立项审批/
    └── ...
```

处理内容：替换「审计期间 / 编制·日期 / 复核·日期」并统一字体 Arial；在 GC 底稿的「支持文档」右侧列填入文档清单；
在 ELC 的「测试汇总表」按 C 列编码匹配文档，填入 J 列（名称）和 K 列（索引）；处理「IT 审计发现汇总表」与「IT 复杂性情况表」。

> 缺少 `FileName.xlsx` / `ITaudit_v1/`、或 `FileName.xlsx` 里没有 `folder` 列时，**会明确报错退出**（不会再静默生成空底稿）。

---

## 作为库使用

```python
from forwork import get_catalogue
from forwork import XlsxReader, random_sample_from_reader, write_sample_xlsx

# 文件清单
df = get_catalogue("D:/Work/底稿", "xlsx", export=False)

# 纯标准库读取 + 可复现抽样
reader = XlsxReader("数据.xlsx")
rows = random_sample_from_reader(reader, "Sheet1", n=10, seed=42)
```

其它子模块也可以按需导入：

```python
from forwork.datamerge import find_files, load_files, write_all
from forwork.repics import reduce_size
from forwork.inspics import insert_images_to_excel    # 需要 xlwings（Windows）
from forwork.itmwp import myWorkingPaper              # 需要 xlwings（Windows）
```

在自己的项目里引用这个库（可编辑，改代码立即生效）：

```bash
cd ~/我的项目 && uv add --editable ~/projects/forwork
```

---

## 常见问题

**Q：`forwork merge` 报错怎么都是中文，不是 traceback？**
这是刻意设计的：所有错误路径统一走 `_fail()`，打印可读的中文提示并返回退出码 1，而不是抛异常堆栈。

**Q：为什么 `merge -t xls` 不支持？**
`.xls` 是旧版二进制格式，pandas 读取它需要额外的 `xlrd` 依赖（且已停止维护）；写出 `.xls` 也不被 openpyxl 支持。
请先在 Excel 里另存为 `.xlsx`。`sample` / `inspics` 同理只支持 `.xlsx`。

**Q：`forwork catalogue xlsx` 列出的是源码目录里的文件？**
0.2.0 起所有工具的默认目录都改成了**当前工作目录**，不再是脚本所在目录。
用 `-d` 可以显式指定任意目录。

**Q：WSL 里能不能跑 `inspics` / `itmwp`？**
不能。它们通过 xlwings 调用**本机** Excel，Excel 装在 Windows 上，WSL 里起不来 COM。
在 WSL 执行会得到明确提示（而不是异常堆栈）。这两个命令请到 Windows 侧运行。

**Q：Excel 进程残留了 / 底稿生成到一半崩了？**
`inspics` 现在用 `try/finally` 保证异常时也会关闭 Excel 进程；`itmwp` 的 `with` 语句同样保证退出时释放。

**Q：怎么确认改的代码生效了？**
`uv tool install -e` 是可编辑安装，`.py` 改完立即生效。命令行入口是软链，指向 `~/.local/share/uv/tools/forwork/bin/forwork`。

---

## 项目结构

```
forwork/
├── pyproject.toml        # 打包配置、依赖、命令行入口 (forwork = forwork.cli:main)
├── uv.lock               # 锁文件
├── .python-version       # 开发用 Python 版本 (3.15)
└── src/forwork/
    ├── __init__.py       # 包入口，暴露公共 API 与 __version__
    ├── __main__.py       # 支持 python -m forwork
    ├── cli.py            # 统一子命令调度器
    ├── _excel.py         # xlwings 延迟导入 + 平台检查（仅 Windows 可用）
    ├── catalogue.py      # 文件清单
    ├── datamerge.py      # 数据合并
    ├── sampling.py       # 随机抽样（纯标准库）
    ├── inspics.py        # 图片插入 Excel
    ├── repics.py         # 图片缩放
    └── itmwp.py          # IT 审计工作底稿生成
```

每个工具模块都导出三个约定函数，便于统一调度与单独测试：

- `configure_parser(parser)` —— 向解析器注册参数
- `run(args)` —— 执行主体，返回退出码
- `main(argv=None)` —— 独立运行入口（内部自建解析器）

---

## 开发指南

```bash
# 安装依赖 + 创建 .venv
uv sync

# 跑起来（用项目虚拟环境）
uv run forwork --help
uv run forwork merge -d ./data -t csv --dry-run

# 代码检查（ruff）
uvx ruff check src/
uvx ruff format src/

# 重新注册全局命令（改了 pyproject 依赖后需要）
uv tool install -e . --reinstall

# 提交
git add -A && git commit -m "..." && git push
```

**约定**：

- 每个模块头部写明【调用方式】（WSL / Windows / 库调用）、【平台要求】、【路径规则】、【默认值】。
- 函数的 docstring 用 `Args / Returns / Raises` 结构，描述**实际怎么调用**而不是内部实现。
- 错误一律走友好提示 + 返回退出码，不要让用户看 Python traceback。
- 需要 Excel 的功能走 `forwork._excel.require_xlwings()`，保证在没有 xlwings 的机器上 `--help` 仍可用。
