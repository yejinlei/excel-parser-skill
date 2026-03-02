---
name: excel-parser
description: 支持多种Excel格式的内容提取技能，使用calamine库，可处理.xls、.xlsx、.xlsm等格式
version: 1.0.0
author: Excel Parser Skill Team
license: MIT
tags:
  - excel
  - spreadsheet
  - data-extraction
  - calamine
  - xls
  - xlsx
  - xlsm
  - data-analysis
---

# Excel Parser Skill

Excel Parser技能用于从Excel文件中提取内容，支持多种Excel格式。该技能使用calamine库，能够处理各种Excel文件格式，包括旧版的.xls和新版的.xlsx、.xlsm等。

## 功能特性

- 支持多种Excel格式：.xls、.xlsx、.xlsm、.xltx、.xltm
- 提取所有工作表的内容
- 保持表格结构和数据格式
- 自动安装缺失的依赖
- 输出结构化数据和文本格式结果
- 支持处理大型Excel文件
- 跨平台兼容

## 安装

### 依赖要求

```bash
pip install python-dotenv
```

### 核心依赖（会自动安装）

```bash
pip install calamine
```

## 环境变量配置

1. 复制 `.env.example` 文件并重命名为 `.env`
2. 根据需要配置以下选项：

```env
# Excel解析器配置
# 最大行数限制，默认100行
EXCEL_MAX_ROWS=100

# 是否保留空行，默认false
EXCEL_KEEP_EMPTY_ROWS=false
```

## 快速开始

### 基本使用

```python
# 导入Excel解析器
from excel_parser import ExcelParser

# 创建解析器实例
parser = ExcelParser()

# 执行Excel解析
result = parser.parse_excel('path/to/your/file.xlsx')

# 获取解析结果
print(f"解析完成，共 {result['sheet_count']} 个工作表")
print(f"总单元格数: {result['total_cells']}")
print(f"使用引擎: {result['engine']}")

# 查看工作表数据
for sheet in result['sheets']:
    print(f"\n工作表: {sheet['name']}")
    print(f"行数: {sheet['row_count']}, 列数: {sheet['column_count']}")
    print("前5行数据:")
    for row in sheet['rows'][:5]:
        print(row)
```

### 转换为文本格式

```python
# 导入Excel解析器
from excel_parser import ExcelParser

# 创建解析器实例
parser = ExcelParser()

# 执行Excel解析并转换为文本
text = parser.parse_excel_to_text('path/to/your/file.xls')

# 打印文本结果
print(text)
```

### 使用主函数

```python
# 导入处理函数
from excel_parser import process_excel

# 处理Excel文件
result = process_excel('path/to/your/file.xlsm')

# 获取结果
print(f"解析完成，共 {result['sheet_count']} 个工作表")
print(f"总单元格数: {result['total_cells']}")
print("\n文本结果:")
print(result['text'])
```

### 命令行使用

```bash
# 处理Excel文件
python excel_parser.py your_file.xlsx
```

## 进阶使用示例

### 批量处理多个Excel文件

```python
import os
from excel_parser import process_excel

# 批量处理目录中的所有Excel文件
excel_dir = "path/to/excel/files"
output_dir = "path/to/output"
os.makedirs(output_dir, exist_ok=True)

for excel_file in os.listdir(excel_dir):
    if excel_file.endswith(('.xls', '.xlsx', '.xlsm', '.xltx', '.xltm')):
        excel_path = os.path.join(excel_dir, excel_file)
        output_path = os.path.join(output_dir, f"{os.path.splitext(excel_file)[0]}.txt")
        
        print(f"处理文件: {excel_file}")
        try:
            result = process_excel(excel_path)
            
            # 保存解析结果到文本文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(result['text'])
            
            print(f"处理完成，结果已保存到: {output_path}")
        except Exception as e:
            print(f"处理失败: {e}")
```

### 提取特定工作表数据

```python
from excel_parser import ExcelParser

def extract_specific_sheet(excel_path, sheet_name):
    """提取特定工作表的数据"""
    parser = ExcelParser()
    result = parser.parse_excel(excel_path)
    
    for sheet in result['sheets']:
        if sheet['name'] == sheet_name:
            return sheet
    
    return None

# 使用示例
sheet_data = extract_specific_sheet('path/to/your/file.xlsx', 'Sheet1')
if sheet_data:
    print(f"工作表: {sheet_data['name']}")
    print(f"行数: {sheet_data['row_count']}, 列数: {sheet_data['column_count']}")
    print("数据:")
    for row in sheet_data['rows']:
        print(row)
else:
    print("未找到指定的工作表")
```

## 支持的文件格式

- **Excel 97-2003**: .xls
- **Excel 2007+**: .xlsx, .xlsm
- **Excel模板**: .xltx, .xltm

## 输出格式

### 结构化数据格式

```python
{
    "text": "解析的文本内容",
    "sheets": [
        {
            "name": "Sheet1",
            "rows": [["A1", "B1", "C1"], ["A2", "B2", "C2"]],
            "row_count": 2,
            "column_count": 3
        }
    ],
    "sheet_count": 1,
    "total_cells": 6,
    "engine": "calamine"
}
```

### 文本格式

```
=== Excel文件: example.xlsx ===
工作表数量: 2
总单元格数: 20

--- 工作表: Sheet1 ---
行数: 5, 列数: 2

1: 姓名	年龄
2: 张三	25
3: 李四	30
4: 王五	28

--- 工作表: Sheet2 ---
行数: 3, 列数: 3

1: 产品	价格	数量
2: 苹果	5.5	100
3: 香蕉	3.2	200
```

## 使用场景

- 提取Excel表格中的数据进行分析
- 批量处理Excel文件并转换为文本格式
- 从旧版.xls文件中提取数据
- 自动化处理Excel报表
- 数据迁移和转换
- 内容审核和数据验证

## 注意事项

1. **性能考虑**：
   - 对于大型Excel文件，解析速度可能会有所下降
   - 默认只输出前100行数据，可通过环境变量调整

2. **数据格式**：
   - 所有单元格值都会被转换为字符串
   - 空单元格会被转换为空字符串

3. **文件大小**：
   - 支持处理大型Excel文件，但建议对超大文件进行分批处理

## 触发使用的提示词

在与 AI IDE 中的助手交互时，您可以使用以下提示词来触发Excel解析技能：

### 📍 触发Excel解析的提示词
- "帮我解析这个Excel文件"
- "提取Excel表格中的数据"
- "处理这个.xls文件"
- "将Excel转换为文本"
- "分析Excel中的数据"
- "提取Excel工作表内容"

### 📍 示例对话

**示例 1：基本解析**
```
用户：帮我解析这个Excel文件，提取所有数据
助手：好的，我将使用Excel Parser技能为您处理。请提供Excel文件路径。
```

**示例 2：转换为文本**
```
用户：将这个.xls文件转换为文本格式
助手：理解，我将使用Excel Parser技能将文件转换为文本。请提供文件路径。
```

**示例 3：批量处理**
```
用户：批量处理目录中的所有Excel文件
助手：我将使用Excel Parser技能批量处理。请提供目录路径。
```

### 🔧 技术实现

当 AI 助手接收到这些提示词时，会：

1. 解析用户意图，确定要处理的Excel文件
2. 调用 ExcelParser 类或 process_excel 函数
3. 执行解析并返回结果

### 🎯 最佳实践

- **明确文件路径**：提供完整的Excel文件路径
- **说明格式**：如果是旧版.xls文件，最好说明一下
- **指定需求**：明确您需要提取的具体内容或格式
- **测试不同文件**：对于不同格式的Excel文件，可能需要调整解析参数

通过使用这些提示词，您可以在与 AI IDE 交互时灵活使用Excel解析技能，快速提取和处理Excel数据

## 故障排除

### 常见问题及解决方案

1. **Calamine依赖安装失败**
   - 问题：`ModuleNotFoundError: No module named 'calamine'`
   - 解决方案：手动安装Calamine依赖：`pip install calamine`

2. **文件格式不支持**
   - 问题：`Excel解析失败: Unsupported file format`
   - 解决方案：确保文件是有效的Excel格式（.xls, .xlsx, .xlsm等）

3. **文件读取权限**
   - 问题：`Excel解析失败: Permission denied`
   - 解决方案：确保您有读取该文件的权限

4. **文件损坏**
   - 问题：`Excel解析失败: File is corrupted`
   - 解决方案：检查Excel文件是否损坏，尝试打开验证

5. **内存不足**
   - 问题：`Excel解析失败: MemoryError`
   - 解决方案：对于大型Excel文件，考虑分批处理或增加内存

## 许可证

MIT License - 详见 [LICENSE.txt](LICENSE.txt)
