# Excel/CSV to Markdown Converter

一个简单而强大的工具，用于将Excel和CSV文件转换为Markdown表格格式。

## 功能特点

- 支持Excel (.xlsx, .xls)和CSV文件格式
- 支持Excel文件中的多个工作表
- 支持批量处理整个目录中的所有Excel和CSV文件
- 提供命令行界面和图形用户界面
- 可以预览生成的Markdown表格
- 可以将结果复制到剪贴板或保存到文件
- 可以选择是否包含表头

## 安装

### 前提条件

- Python 3.7+
- pip (Python包管理器)

### 安装步骤

1. 克隆或下载此仓库
```bash
git clone https://github.com/foxstarx2beijing/excel_to_markdown
cd excel_to_markdown
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法

### 命令行界面

```bash
# 基本用法 - 转换单个文件
python excel_to_md.py input_file.xlsx

# 指定输出文件
python excel_to_md.py input_file.xlsx -o output_file.md

# 指定Excel工作表
python excel_to_md.py input_file.xlsx -s "Sheet1"

# 预览输出
python excel_to_md.py input_file.xlsx -p

# 处理整个目录中的所有Excel/CSV文件
python excel_to_md.py /path/to/directory --directory -o /path/to/output_dir

# 递归处理目录及其子目录中的所有Excel/CSV文件
python excel_to_md.py /path/to/directory --directory --recursive
```

命令行参数:
- `input_path`: 输入的Excel/CSV文件路径或目录路径（使用--directory选项时）
- `-o, --output`: 输出的Markdown文件路径或目录路径
- `-s, --sheet`: 要转换的Excel工作表名称
- `-p, --preview`: 在终端中预览输出
- `-d, --directory`: 将input_path视为目录，处理目录中的所有Excel/CSV文件
- `-r, --recursive`: 与--directory一起使用，递归处理子目录

### 图形用户界面

运行GUI应用程序:
```bash
python gui.py
```

1. 点击"Browse"按钮选择输入文件
2. 如果是Excel文件，可以从下拉列表中选择工作表
3. 选择是否包含表头
4. 点击"Browse"按钮选择输出文件位置
5. 在预览区域中查看生成的Markdown表格
6. 点击"Copy to Clipboard"复制到剪贴板
7. 点击"Convert & Save"保存到文件

### 作为Python模块运行

```bash
# CLI版本
python -m excel_to_markdown

# GUI版本
python -m excel_to_markdown --gui

# 处理目录
python -m excel_to_markdown /path/to/directory --directory
```

## 程序库使用方法

你也可以在自己的Python项目中使用此程序库:

```python
from excel_to_md import convert_excel_to_markdown, convert_directory

# 转换单个文件并获取Markdown文本
markdown = convert_excel_to_markdown("input.xlsx", sheet_name="Sheet1")

# 转换单个文件并保存到文件
convert_excel_to_markdown("input.xlsx", "output.md")

# 转换整个目录中的文件
convert_directory("input_directory", "output_directory", recursive=True)
```

## 示例

### 输入: Excel表格
| Name | Age | City |
|------|-----|------|
| John | 30  | New York |
| Alice | 25 | London |
| Bob | 35 | Paris |

### 输出: Markdown表格
```markdown
| Name | Age | City |
| --- | --- | --- |
| John | 30 | New York |
| Alice | 25 | London |
| Bob | 35 | Paris |
```

## 依赖包

- pandas: 数据处理
- openpyxl/xlrd: Excel文件读取
- click: 命令行界面
- rich: 终端美化
- tkinter: 图形用户界面

## 许可证

MIT

## 作者

foxstarx 