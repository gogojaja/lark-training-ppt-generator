# 技能包总索引

**生成时间**: 2026-08-13
**版本**: v1.0.0

## 概述

本索引提供了所有技能包的快速导航和基本信息，帮助用户快速找到需要的技能包。

## 技能包列表

### 1. cover-skill

**技能名称**: 封面页技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于创建专业PPT封面页，包含标题、副标题、作者、日期等信息

**功能列表**:
- 创建标准封面页
- 支持自定义样式
- 支持多种配色方案

**使用方法**:
```python
from cover_skill import create_cover_page

slide = create_cover_page(
    pres,
    title='个人综合签约培训',
    subtitle='操作流程与注意事项',
    author='培训部',
    date='2026-08-13'
)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 2. toc-skill

**技能名称**: 目录页技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于创建专业PPT目录页，展示培训内容的结构

**功能列表**:
- 创建标准目录页
- 支持多列布局
- 支持页码显示

**使用方法**:
```python
from toc_skill import create_toc_page

items = [
    '个人综合签约概述 - 第1-2页',
    '签约流程详解 - 第3-8页',
    '注意事项 - 第9-12页',
    '常见问题 - 第13-15页'
]
slide = create_toc_page(pres, '目录', items)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 3. scene-description-skill

**技能名称**: 场景描述技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于创建业务场景描述页，展示业务场景和关键要素

**功能列表**:
- 创建场景描述页
- 支持角色信息展示
- 支持环境信息展示

**使用方法**:
```python
from scene_description_skill import create_scene_description

slide = create_scene_description(
    pres,
    title='签约场景',
    scenario='客户经理与客户在办公室进行签约',
    characters='客户经理：负责引导和解释\n客户：提交申请和确认信息',
    environment='办公室环境，安静私密'
)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 4. business-rules-skill

**技能名称**: 业务规则技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于展示业务规则和约束条件

**功能列表**:
- 创建业务规则页
- 支持规则列表展示
- 支持规则优先级标注

**使用方法**:
```python
from business_rules_skill import create_business_rules_page

rules = [
    '规则1：所有签约必须经过风控审核',
    '规则2：签约材料必须在3个工作日内完成审核',
    '规则3：合同必须由法务部门审核通过'
]
slide = create_business_rules_page(pres, '业务规则', rules)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 5. operation-steps-skill

**技能名称**: 操作步骤技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于展示操作步骤和流程

**功能列表**:
- 创建操作步骤页
- 支持步骤编号
- 支持步骤说明

**使用方法**:
```python
from operation_steps_skill import create_operation_steps_page

steps = [
    '步骤1：提交签约申请',
    '步骤2：准备相关材料',
    '步骤3：等待审核',
    '步骤4：签订合同'
]
slide = create_operation_steps_page(pres, '操作步骤', steps)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 6. key-points-skill

**技能名称**: 关键点技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于展示培训的关键点和注意事项

**功能列表**:
- 创建关键点页
- 支持要点列表
- 支持重点标注

**使用方法**:
```python
from key_points_skill import create_key_points_page

points = [
    '关键点1：签约前需确认客户身份',
    '关键点2：材料必须齐全',
    '关键点3：合同条款需仔细阅读',
    '关键点4：签约后需及时归档'
]
slide = create_key_points_page(pres, '关键点', points)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 7. faq-skill

**技能名称**: 常见问题技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于展示常见问题和解答

**功能列表**:
- 创建FAQ页
- 支持问答列表
- 支持问题分类

**使用方法**:
```python
from faq_skill import create_faq_page

faqs = [
    '问题1：签约需要多长时间？\n解答：通常需要3-5个工作日',
    '问题2：材料不全可以签约吗？\n解答：不可以，材料必须齐全',
    '问题3：可以在线签约吗？\n解答：目前只支持线下签约'
]
slide = create_faq_page(pres, '常见问题', faqs)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 8. ppt-framework

**技能名称**: PPT框架技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: PPT生成的基础框架，提供页面生成、样式管理、布局系统

**功能列表**:
- 页面生成器
- 样式管理系统
- 布局系统
- 垂直节奏系统

**使用方法**:
```python
from ppt_framework import generate_page, manage_style

# 生成页面
slide = generate_page(pres, 'cover', config)

# 管理样式
manage_style(pres, theme='blue')

# 布局系统
layout_system(pres, layout='flex')

# 垂直节奏
vertical_rhythm(pres, base=16, ratio=1.5)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 所有技能包

**依赖技能**: 无

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 9. style-brief-skill

**技能名称**: 风格简报技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于创建风格简报，展示PPT的视觉风格

**功能列表**:
- 创建风格简报
- 支持配色方案展示
- 支持字体展示
- 支持组件展示

**使用方法**:
```python
from style_brief_skill import create_style_brief

style_brief = {
    'primary_color': '#1F3864',
    'secondary_color': '#2E75B6',
    'font_family': 'Microsoft YaHei',
    'font_size': 16,
    'spacing': 16
}
slide = create_style_brief(pres, '风格简报', style_brief)
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 所有技能包

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 10. table-skill

**技能名称**: 表格技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于创建专业表格，支持多种表格样式

**功能列表**:
- 创建标准表格
- 创建斑马纹表格
- 创建图标表格
- 支持多种配色方案

**使用方法**:
```python
from table_skill import create_table

columns = ['步骤', '操作', '时间', '负责人']
rows = [
    ['1', '提交申请', '1-2天', '客户经理'],
    ['2', '审核材料', '2-3天', '风控专员'],
    ['3', '签订合同', '1天', '法务部门']
]
slide = create_table(pres, '签约流程', columns, rows, table_type='zebra')
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 11. document-processing

**技能名称**: 文档处理技能包
**技能版本**: v1.1.0
**发布日期**: 2026-08-07
**技能描述**: 用于处理Word文档，包括章节拆分、结构分析、格式转换

**功能列表**:
- Word文档章节拆分
- 文档结构分析
- 批量拆分
- 格式转换

**使用方法**:
```bash
# 拆分所有level=1的章节
py -3 tools/split_docx_by_level.py input.docx output_dir

# 只拆分第3个level=1章节
py -3 tools/split_docx_by_level.py input.docx output_dir 1 3

# 按level=2拆分
py -3 tools/split_docx_by_level.py input.docx output_dir 2
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: 无

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

### 12. flowchart-skill

**技能名称**: 流程图技能包
**技能版本**: v1.0.0
**发布日期**: 2026-08-13
**技能描述**: 用于创建流程图，展示业务流程和决策路径

**功能列表**:
- 创建标准流程图
- 创建泳道图
- 创建循环流程图
- 支持多种流程图类型

**使用方法**:
```python
from flowchart_skill import create_flowchart

nodes = ['开始', '签约申请', '材料审核', '合同签订', '完成']
slide = create_flowchart(pres, '个人综合签约流程', nodes, [])
```

**目录结构**:
- templates/ - 模板文件
- styles/ - 样式定义
- examples/ - 示例代码

**相关技能**: 无

**依赖技能**: ppt-framework

**文档**:
- SKILL_INDEX.md
- README.md
- SKILL.md

---

## 技能包分类

### 页面生成
- cover-skill: 封面页
- toc-skill: 目录页
- scene-description-skill: 场景描述

### 内容展示
- business-rules-skill: 业务规则
- operation-steps-skill: 操作步骤
- key-points-skill: 关键点
- faq-skill: 常见问题

### 数据可视化
- table-skill: 表格
- flowchart-skill: 流程图

### 基础框架
- ppt-framework: PPT框架
- style-brief-skill: 风格简报

### 文档处理
- document-processing: 文档处理

## 依赖关系

```
ppt-framework (基础框架)
    ├── cover-skill
    ├── toc-skill
    ├── scene-description-skill
    ├── business-rules-skill
    ├── operation-steps-skill
    ├── key-points-skill
    ├── faq-skill
    ├── table-skill
    ├── flowchart-skill
    └── style-brief-skill

document-processing (独立)
```

## 使用指南

### 快速开始

1. **选择技能包**: 根据需要选择合适的技能包
2. **查看示例**: 在 `examples/` 目录中查看使用示例
3. **参考文档**: 阅读 `README.md` 和 `SKILL.md`
4. **应用代码**: 将示例代码应用到项目中

### 技能包组合使用

```python
from ppt_framework import Presentation

# 创建演示文稿
pres = Presentation()

# 添加封面页
from cover_skill import create_cover_page
create_cover_page(pres, '个人综合签约培训', '操作流程与注意事项', '培训部', '2026-08-13')

# 添加目录页
from toc_skill import create_toc_page
toc_items = [
    '个人综合签约概述 - 第1-2页',
    '签约流程详解 - 第3-8页',
    '注意事项 - 第9-12页',
    '常见问题 - 第13-15页'
]
create_toc_page(pres, '目录', toc_items)

# 添加操作步骤页
from operation_steps_skill import create_operation_steps_page
steps = [
    '步骤1：提交签约申请',
    '步骤2：准备相关材料',
    '步骤3：等待审核',
    '步骤4：签订合同'
]
create_operation_steps_page(pres, '操作步骤', steps)

# 保存演示文稿
pres.save('output/training.pptx')
```

## 更新日志

### v1.0.0 (2026-08-13)
- 初始版本
- 创建12个技能包
- 完成标准化结构
- 添加基础模板和示例

## 联系方式

如有问题或建议，请提交 Issue 或 Pull Request。

## 许可证

MIT License
