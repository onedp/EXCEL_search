# Excel 内容查询工具

这是一个用于在指定文件夹中的全部Excel文件中搜索内容的GUI工具。它支持中英文语言切换，允许用户在两种语言之间切换。

## 功能

- 在任何一个文件夹的所有Excel文件的所有工作表中搜索内容。
- 支持 `.xlsx` 和 `.xls` 文件格式。
- 多语言支持（中英文）。
- 显示搜索结果，并可选择打开选定的文件。
- 记录错误并在单独的日志窗口中显示。

## 安装

### 先决条件

- Python 3.x
- `pandas` 库
- `openpyxl` 库
- `tkinter` 库（通常包含在Python中）

### 安装依赖项

```bash
pip install pandas openpyxl