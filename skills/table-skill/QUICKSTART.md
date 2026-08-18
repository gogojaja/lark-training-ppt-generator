# table-skill 快速开始

## 1. 提取表格数据

从 Word 文档提取表格数据：

```bash
py -3 extract_table_data.py
```

这将提取"预约查询.docx"的表格数据并保存为 JSON 文件。

## 2. 生成表格页

使用 table-skill 启动脚本生成表格页：

```bash
py -3 skills/table-skill/table_skill.py "生成产物/表格/预约查询.json" zebra blue "我的表格.pptx"
```

## 3. 参数说明

- **json_file**：JSON 数据文件路径（必需）
- **template**：模板类型（可选）
  - `standard` - 标准表格
  - `zebra` - 斑马纹表格
  - `icon` - 带图标表格
- **preset**：配色方案（可选）
  - `blue` - 专业蓝
  - `gray` - 商务灰
- **output_file**：输出文件名（可选）

## 4. 实际案例

### 案例 1：预约查询表格

```bash
py -3 skills/table-skill/table_skill.py "生成产物/表格/预约查询.json" zebra blue "预约查询表格.pptx"
```

### 案例 2：标准表格（专业蓝）

```bash
py -3 skills/table-skill/table_skill.py "生成产物/表格/预约查询.json" standard blue
```

### 案例 3：斑马纹表格（商务灰）

```bash
py -3 skills/table-skill/table_skill.py "生成产物/表格/预约查询.json" zebra gray
```

## 5. JSON 数据格式

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

## 6. 输出示例

生成的表格页包含：
- 表格标题
- 表头（深色背景）
- 数据行
- 斑马纹效果（可选）
- 页脚

## 7. 文件位置

- **启动脚本**：`skills/table-skill/table_skill.py`
- **生成脚本**：`生成脚本/gen_table.py`
- **提取脚本**：`extract_table_data.py`
- **JSON 数据**：`生成产物/表格/预约查询.json`
- **输出文件**：自定义（默认：表格.pptx）
