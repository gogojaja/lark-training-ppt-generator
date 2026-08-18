# table-skill 使用说明

## 概述

table-skill 是一个用于生成专业 PPT 表格页的技能，支持多种模板和配色方案。

## 功能特性

- **多种模板**：标准表格、斑马纹表格、带图标表格
- **多种配色**：专业蓝、商务灰
- **自动提取**：从 Word 文档自动提取表格数据
- **灵活配置**：支持自定义标题和输出路径

## 使用方法

### 1. 提取表格数据

从 Word 文档提取表格数据并保存为 JSON 格式：

```bash
python extract_table_data.py
```

这将提取文档中的第一个表格并保存为 JSON 文件。

### 2. 生成表格页

使用提取的 JSON 数据生成 PPT 表格页：

```bash
python gen_table.py <json_file> [template] [preset] [output_file]
```

**参数说明：**

- `json_file`：JSON 数据文件路径（必需）
- `template`：模板类型（可选，默认：standard）
  - `standard`：标准表格
  - `zebra`：斑马纹表格
  - `icon`：带图标表格
- `preset`：配色方案（可选，默认：blue）
  - `blue`：专业蓝
  - `gray`：商务灰
- `output_file`：输出文件名（可选，默认：表格.pptx）

**示例：**

```bash
# 使用默认设置生成表格页
python gen_table.py "生成产物/表格/预约查询.json"

# 使用斑马纹模板和商务灰配色
python gen_table.py "生成产物/表格/预约查询.json" zebra gray

# 自定义输出文件名
python gen_table.py "生成产物/表格/预约查询.json" standard blue "我的表格.pptx"
```

### 3. 批量生成

可以创建批处理脚本批量生成多个表格页：

```bash
# 生成多个表格页
python gen_table.py "表格1.json" zebra blue "表格1.pptx"
python gen_table.py "表格2.json" standard gray "表格2.pptx"
python gen_table.py "表格3.json" icon blue "表格3.pptx"
```

## JSON 数据格式

表格数据需要符合以下 JSON 格式：

```json
{
  "title": "表格标题",
  "headers": ["列1", "列2", "列3"],
  "rows": [
    ["数据1", "数据2", "数据3"],
    ["数据4", "数据5", "数据6"]
  ]
}
```

## 配色方案

### 专业蓝（blue）

- 表头背景：#1F3864（深蓝）
- 表头文字：#FFFFFF（白色）
- 数据行背景：#FFFFFF（白色）
- 斑马纹背景：#F8FAFC（浅灰）
- 边框颜色：#E0E0E0（浅灰）
- 强调色：#2E75B6（中蓝）

### 商务灰（gray）

- 表头背景：#404040（深灰）
- 表头文字：#FFFFFF（白色）
- 数据行背景：#FFFFFF（白色）
- 斑马纹背景：#F5F5F5（浅灰）
- 边框颜色：#E0E0E0（浅灰）
- 强调色：#808080（中灰）

## 模板说明

### 标准表格（standard）

- 简洁清晰的设计
- 适合一般数据展示
- 表头使用深色背景

### 斑马纹表格（zebra）

- 数据行交替使用浅色背景
- 提高数据可读性
- 适合数据较多的场景

### 带图标表格（icon）

- 支持在表格中添加图标
- 适合需要状态标识的场景
- 需要扩展代码支持

## 实际案例

### 案例1：预约查询配置表

从"预约查询.docx"提取的表格数据：

```json
{
  "title": "预约查询 - 高拍仪品牌配置表",
  "headers": [
    "高拍仪品牌",
    "转换值（分辨率后一栏，✔中填值）",
    "分辨率",
    "DPI转换值",
    "色彩",
    "裁剪方式",
    "旋转"
  ],
  "rows": [
    ["升腾", "219", "2592*1944", "200", "彩色", "自动寻边纠偏", "旋转度可自由调配，需保证拍摄影像端正"],
    ["实达", "221", "2592*1944", "200", "彩色", "自动寻边纠偏", ""],
    ["国光", "224", "2592*1944", "200", "彩色", "自动寻边纠偏", ""]
  ]
}
```

生成命令：

```bash
python gen_table.py "生成产物/表格/预约查询.json" zebra blue "预约查询表格.pptx"
```

## 技术实现

### 依赖库

- `python-pptx`：用于创建 PPT 文件
- `python-docx`（可选）：用于从 Word 文档提取表格数据
- `xml.etree.ElementTree`：用于解析 Word 文档的 XML 结构

### 核心函数

- `create_table_slide(pres, data, template, preset, title)`：创建表格页
- `COLOR_SCHEMES`：定义颜色方案
- `TEMPLATE_TYPES`：定义模板类型

## 注意事项

1. JSON 文件必须包含 `title`、`headers`、`rows` 字段
2. 表格列数必须与表头数量一致
3. 行数必须与数据行数量一致
4. 输出文件名不要包含中文（Windows 系统可能有编码问题）
5. 建议使用斑马纹模板来提高数据可读性

## 未来扩展

- 支持更多模板类型（带图标、带状态、带排序等）
- 支持更多配色方案（语义色、自定义色等）
- 支持导出为 Excel
- 支持条件格式（根据数据值改变颜色）
- 支持排序功能

## 版本历史

- **v1.0.0**（2026-08-13）：
  - 首版发布
  - 支持 3 种模板类型
  - 支持 2 种配色方案
  - 自动提取 Word 文档表格数据
