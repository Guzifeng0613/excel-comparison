# Excel / CSV 比较工具

一个基于 Python + Tkinter 的表格文件比对工具，支持图形界面（拖拽操作）与命令行两种使用方式。可按主键逐行比对两个 Excel/CSV 文件，定位单元格级别的差异，自动生成多 Sheet 的 Excel 报告，并生成差异处标红的高亮备份文件。

- 核心版本：**v2.8**
- 支持格式：`.xlsx`、`.xls`、`.xlsm`、`.csv`

---

## 功能特性

- **图形界面 / 命令行双模式**：GUI 支持文件拖拽（需 `tkinterdnd2`），也可点击选择。
- **按主键比对**：指定一列或多列作为主键匹配行；未指定时自动推荐合适的主键列，默认使用第一列。
- **多种比较口径**：精确匹配、忽略空白、包含匹配、正则匹配。
- **文本标准化选项**：忽略大小写、去首尾空格、合并连续空格、移除换行/制表符。
- **数值容差比较**：可配置数值容差（`123` 与 `123.0` 视为相等）。
- **列名模糊匹配**：两表列名不完全一致时，自动按相似度对齐列名（可关闭）。
- **大文件支持**：分块读取、大文件只读模式、多线程并行比较。
- **差异合并**：可指定以某个文件为准，将差异合并后输出新文件。
- **结果可视化**：
  - 生成多 Sheet 的 Excel 比较报告；
  - 生成高亮备份文件，差异单元格直接标红，整行标注“仅存在于某一文件”等。

---

## 项目结构

```
excel-comparison/
├── excel_compare.py      # 核心模块（v2.8）：比较逻辑 + 内置 GUI + CLI 入口
├── gui_compare.py        # 早期独立拖拽界面（v2.1，旧版，保留备查）
├── spec_file.spec        # PyInstaller 打包配置（对应旧版 gui_compare.py）
├── MEMORY.md             # 开发过程记录与测试笔记
├── README.md             # 项目说明（本文件）
└── .gitignore            # Git 忽略规则
```

> 说明：当前主程序为 `excel_compare.py`，其内部已集成完整的图形界面与命令行。
> `gui_compare.py` 为早期界面版本，依赖旧的 `compare_excel_files` 接口，与 v2.8 核心模块不再兼容，仅作留存参考。

---

## 环境要求

- Python 3.8 及以上版本（开发环境为 Python 3.14 / Windows）
- Tkinter（Windows 官方 Python 安装包默认自带）

### 依赖安装

核心依赖：

```bash
pip install pandas numpy openpyxl
```

可选依赖：

```bash
# 文件拖拽支持（不安装时 GUI 仍可通过点击选择文件）
pip install tkinterdnd2

# 读取旧版 .xls 文件
pip install xlrd
```

一键安装：

```bash
pip install pandas numpy openpyxl tkinterdnd2 xlrd
```

---

## 快速开始

### 方式一：图形界面

不带任何参数运行（或显式加 `--gui`）即可打开图形界面：

```bash
python excel_compare.py
# 或
python excel_compare.py --gui
```

操作步骤：

1. 将两个待比较的文件分别拖入（或点击选择）“文件 1 / 文件 2”区域；如为多工作表 Excel，可填写要比对的工作表名称（CSV 留空）。
2. 在“主键列”中填写用于匹配行的列名（可添加多列）；留空则自动推荐 / 使用第一列。
3. 根据需要设置比较类型、标准化选项、数值容差、列名模糊匹配、大文件模式、并行等。
4. 点击“开始比较”，完成后会生成 Excel 报告与高亮备份文件。

### 方式二：命令行

```bash
python excel_compare.py <文件1> <文件2> [选项]
```

示例：

```bash
# 基本比对（自动选择主键，报告生成在当前目录）
python excel_compare.py a.xlsx b.xlsx

# 指定主键列与报告输出路径
python excel_compare.py a.xlsx b.xlsx -k 学号 -o 报告.xlsx

# 使用多列联合主键（-k 可重复出现）
python excel_compare.py a.xlsx b.xlsx -k 年级 -k 班级 -k 姓名

# 忽略大小写、启用并行比较、设置数值容差
python excel_compare.py a.xlsx b.xlsx --normalize-case --parallel --numeric-tolerance 0.01

# 比对 CSV / 指定工作表
python excel_compare.py a.csv b.csv
python excel_compare.py a.xlsx b.xlsx --sheet1 Sheet1 --sheet2 Sheet1

# 大文件只读模式、关闭列名模糊匹配
python excel_compare.py big1.xlsx big2.xlsx --large-file --no-fuzzy

# 差异合并：以文件1为准（--merge 1）或以文件2为准（--merge 2）
python excel_compare.py a.xlsx b.xlsx --merge 1 --merge-output 合并结果.xlsx
```

### 命令行参数说明

| 参数 | 说明 |
| --- | --- |
| `file1`、`file2` | 两个待比较文件的路径（位置参数） |
| `-o, --output` | 比较报告的输出路径 |
| `--sheet1`、`--sheet2` | 指定两个 Excel 文件的工作表名称 |
| `-k, --key` | 主键列名，可重复使用以指定多列联合主键 |
| `--normalize-case` | 比较时忽略大小写 |
| `--keep-whitespace` | 保留空白（默认会去除/合并空白与换行） |
| `--large-file` | 大文件只读模式 |
| `--no-fuzzy` | 禁用列名模糊匹配 |
| `--parallel` | 启用多线程并行比较 |
| `--numeric-tolerance` | 数值比较容差，默认 `1e-9` |
| `--gui` | 强制启动图形界面 |
| `--merge` | 合并差异：`1` 表示以文件 1 为准，`2` 表示以文件 2 为准 |
| `--merge-output` | 差异合并结果的输出路径 |

---

## 比较口径说明

- **精确匹配**：标准化后的值完全相等即视为一致；数值按容差比较（如 `123` 与 `123.0`）。
- **忽略空白**：在精确匹配基础上忽略首尾及多余空白。
- **包含匹配**：判断一个值是否包含另一个值。
- **正则匹配**：使用正则表达式进行匹配判断。

配合“忽略大小写 / 去首尾空格 / 合并空格 / 移除换行”等标准化选项，可消除常见的格式性差异，让比对结果聚焦于真实内容差异。

---

## 输出结果说明

### 1. Excel 比较报告（多 Sheet）

| Sheet | 内容 |
| --- | --- |
| **汇总** | 两个文件的行数、主键列、总主键数、匹配/差异行数、仅存在于各文件的行数及匹配率 |
| **列名映射** | 两表列名模糊匹配的对应关系及相似度（仅在发生列名对齐时生成） |
| **差异明细** | 每条差异的主键值、差异列、差异类型、双方行号/列位置、原始值与标准化值、Excel 单元格位置 |
| **差异列统计** | 各列差异数量及占比，按数量降序排列 |
| **仅文件 1 / 仅文件 2** | 只在其中一个文件中出现的主键行 |
| **颜色图例** | 高亮颜色与差异类型的对应说明 |

### 2. 高亮备份文件

在原文件基础上生成高亮副本，不改动原始文件：

- **红色单元格**：主键匹配但该单元格值不同；
- **橙色整行**：该行仅存在于文件 1；
- **浅蓝色整行**：该行仅存在于文件 2；
- **紫色整行**：主键存在但排序不同 / 无法对应。

---

## 打包为可执行文件（Windows）

使用 PyInstaller 将主程序打包为 exe：

```bash
pip install pyinstaller

# 以当前核心程序为入口，打包为单文件、无控制台窗口的 GUI 程序
pyinstaller -F -w --name Excel比较工具 excel_compare.py
```

打包产物位于 `dist/` 目录。

> 仓库中的 `spec_file.spec` 是针对早期 `gui_compare.py` 的打包配置；如需使用当前核心程序，建议直接使用上面的命令重新生成。

---

## 备注

- 比对不会修改原始文件，所有结果均写入新生成的报告与高亮备份文件。
- 如拖拽功能不可用，请安装 `tkinterdnd2`，或直接点击界面选择文件。
- 开发记录与测试样例见 [MEMORY.md](MEMORY.md)。
