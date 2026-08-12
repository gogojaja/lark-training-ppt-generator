# Skill: table-skill

# 表格页生成技能

## 1. 元数据

- **技能名称**：table-skill
- **技能版本**：v1.0.0
- **发布日期**：2026-08-12
- **技能定位**：生成专业的PPT表格页，展示结构化数据
- **适用场景**：数据对比、规则说明、配置参数、收费标准

## 2. 设计规范

### 2.1 布局模板

**模板A：标准表格**
```
┌─────────────────────────────────────┐
│  表格标题                            │
├─────────────────────────────────────┤
│  ┌─────┬─────┬─────┬─────┐        │
│  │ 列1 │ 列2 │ 列3 │ 列4 │        │
│  ├─────┼─────┼─────┼─────┤        │
│  │ 数据│ 数据│ 数据│ 数据│        │
│  ├─────┼─────┼─────┼─────┤        │
│  │ 数据│ 数据│ 数据│ 数据│        │
│  └─────┴─────┴─────┴─────┘        │
└─────────────────────────────────────┘
```

**模板B：斑马纹表格**
```
┌─────────────────────────────────────┐
│  表格标题                            │
├─────────────────────────────────────┤
│  ┌─────┬─────┬─────┬─────┐        │
│  │ 列1 │ 列2 │ 列3 │ 列4 │        │
│  ├─────┼─────┼─────┼─────┤        │
│  │ 数据│ 数据│ 数据│ 数据│        │ ← 浅灰背景
│  ├─────┼─────┼─────┼─────┤        │
│  │ 数据│ 数据│ 数据│ 数据│        │ ← 白色背景
│  └─────┴─────┴─────┴─────┘        │
└─────────────────────────────────────┘
```

**模板C：带图标表格**
```
┌─────────────────────────────────────┐
│  表格标题                            │
├─────────────────────────────────────┤
│  ┌─────┬─────┬─────┬─────┐        │
│  │ 图标│ 列2 │ 列3 │ 列4 │        │
│  ├─────┼─────┼─────┼─────┤        │
│  │  ✅ │ 数据│ 数据│ 数据│        │
│  ├─────┼─────┼─────┼─────┤        │
│  │  ❌ │ 数据│ 数据│ 数据│        │
│  └─────┴─────┴─────┴─────┘        │
└─────────────────────────────────────┘
```

### 2.2 配色方案

**专业蓝（默认）**：
- 表头背景：#1F3864（深蓝）
- 表头文字：#FFFFFF（白色）
- 数据行背景：#FFFFFF（白色）
- 斑马纹背景：#F8FAFC（浅灰）
- 边框颜色：#E0E0E0（浅灰）
- 强调色：#2E75B6（中蓝）

**商务灰**：
- 表头背景：#404040（深灰）
- 表头文字：#FFFFFF（白色）
- 数据行背景：#FFFFFF（白色）
- 斑马纹背景：#F5F5F5（浅灰）
- 边框颜色：#E0E0E0（浅灰）
- 强调色：#808080（中灰）

**语义色**：
- 成功/通过：#006100（绿色）
- 警告/注意：#7F6000（黄色）
- 错误/禁止：#C00000（红色）
- 信息/提示：#2E75B6（蓝色）

### 2.3 字体规范

| 元素 | 字号 | 字体 | 颜色 | 对齐 |
|------|------|------|------|------|
| 页面标题 | 28pt | 微软雅黑 | 深灰 | 左对齐 |
| 表头文字 | 14pt | 微软雅黑（加粗） | 白色 | 居中 |
| 数据行文字 | 12pt | 微软雅黑 | 深灰 | 左对齐 |
| 强调文字 | 12pt | 微软雅黑（加粗） | 主色 | 左对齐 |

## 3. 工作流

```
数据表格 → ①选择模板 → ②配置样式 → ③生成PPT
```

### Step 1：收集表格数据

**数据格式**：
```json
{
  "title": "收费标准",
  "headers": ["服务项目", "收费标准", "备注"],
  "rows": [
    ["UKey管理", "50元/个", "一次性收费"],
    ["短信通知", "2元/月", "按月扣费"],
    ["跨行转账", "按笔收费", "详见费率表"]
  ]
}
```

### Step 2：选择模板和配色

**模板选择原则**：
- 标准表格：通用场景
- 斑马纹表格：数据较多时
- 带图标表格：需要状态标识时

**配色选择原则**：
- 专业蓝：技术/业务类
- 商务灰：管理/行政类
- 语义色：需要状态标识时

### Step 3：生成表格页

**生成脚本**：
```bash
# 使用默认模板
py -3 生成脚本/gen_table.py --data "收费标准.json" --out 表格.pptx

# 指定模板和配色
py -3 生成脚本/gen_table.py --data "收费标准.json" --template zebra --preset blue --out 表格.pptx

# 从CSV生成
py -3 生成脚本/gen_table.py --csv "收费标准.csv" --template icon --preset gray --out 表格.pptx
```

## 4. 代码实现

### 4.1 XML结构

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <!-- 标题 -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="title"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="500000" y="200000"/><a:ext cx="11192000" cy="600000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" anchor="ctr"/>
          <a:p>
            <a:r>
              <a:rPr lang="zh-CN" sz="2800" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
              </a:rPr>
              <a:t>表格标题</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      
      <!-- 表格 -->
      <p:tbl>
        <p:tblPr>
          <a:tblW w="11192000" type="dist"/>
          <a:tblBorders>
            <a:top w="9525" val="single" color="E0E0E0"/>
            <a:left w="9525" val="single" color="E0E0E0"/>
            <a:bottom w="9525" val="single" color="E0E0E0"/>
            <a:right w="9525" val="single" color="E0E0E0"/>
            <a:insideH w="9525" val="single" color="E0E0E0"/>
            <a:insideV w="9525" val="single" color="E0E0E0"/>
          </a:tblBorders>
        </p:tblPr>
        <p:tblGrid>
          <p:gridCol w="3730667"/>
          <p:gridCol w="3730667"/>
          <p:gridCol w="3730666"/>
        </p:tblGrid>
        
        <!-- 表头行 -->
        <p:tr>
          <p:trPr><a:trHeight val="457200"/></p:trPr>
          <p:tc>
            <p:tcPr>
              <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
            </p:tcPr>
            <p:txBody>
              <a:bodyPr wrap="square" anchor="ctr"/>
              <a:p>
                <a:pPr algn="ctr"/>
                <a:r>
                  <a:rPr lang="zh-CN" sz="1400" b="1" dirty="0">
                    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                  </a:rPr>
                  <a:t>表头</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:tc>
        </p:tr>
        
        <!-- 数据行 -->
        <p:tr>
          <p:trPr><a:trHeight val="400000"/></p:trPr>
          <p:tc>
            <p:tcPr>
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            </p:tcPr>
            <p:txBody>
              <a:bodyPr wrap="square" anchor="ctr"/>
              <a:p>
                <a:pPr algn="ctr"/>
                <a:r>
                  <a:rPr lang="zh-CN" sz="1200" dirty="0"/>
                  <a:t>数据</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:tc>
        </p:tr>
      </p:tbl>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
```

## 5. 最佳实践

### 5.1 设计原则

1. **对齐美观**：保持表格元素对齐
2. **颜色协调**：使用统一的配色方案
3. **层次清晰**：通过表头、斑马纹区分信息层次
4. **留白适当**：避免表格过于拥挤

### 5.2 常见问题

1. **列宽不均**：合理分配列宽
2. **文字溢出**：控制文本长度，必要时换行
3. **颜色不协调**：使用预设配色方案
4. **数据不清晰**：使用斑马纹或图标增强可读性

### 5.3 优化建议

1. **添加图标**：为状态字段添加图标
2. **条件格式**：根据数据值改变颜色
3. **排序功能**：支持按列排序
4. **导出功能**：支持导出为Excel

## 6. 版本更新日志

### v1.0.0（2026-08-12）
- 首版发布：支持3种表格模板
- 支持3种配色方案
- 支持表头、数据行、斑马纹
