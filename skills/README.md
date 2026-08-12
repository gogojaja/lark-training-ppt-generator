# PPT生成技能库

## 技能概览

本技能库提供从Word文档生成专业PPT的完整解决方案，涵盖封面、目录、流程图、表格、场景说明、业务规则、操作步骤、业务要点、常见问题等所有PPT页面类型。

## 技能列表

### 风格系统技能

| 技能名称 | 版本 | 功能 | 适用场景 |
|----------|------|------|----------|
| [style-brief-skill](style-brief-skill/SKILL.md) | v1.0.0 | 风格简报生成 | 风格选择、气质定义 |

### 基础页面技能

| 技能名称 | 版本 | 功能 | 适用场景 |
|----------|------|------|----------|
| [cover-skill](cover-skill/SKILL.md) | v1.0.0 | 封面页生成 | PPT封面、标题页 |
| [toc-skill](toc-skill/SKILL.md) | v1.0.0 | 目录页生成 | 文档结构、章节导航 |
| [table-skill](table-skill/SKILL.md) | v1.0.0 | 表格页生成 | 数据对比、规则说明 |

### 内容页面技能

| 技能名称 | 版本 | 功能 | 适用场景 |
|----------|------|------|----------|
| [flowchart-skill](flowchart-skill/SKILL.md) | v1.2.0 | 流程图生成 | 业务流程、操作流程 |
| [scene-description-skill](scene-description-skill/SKILL.md) | v1.1.0 | 场景说明页 | 业务背景、功能介绍 |
| [business-rules-skill](business-rules-skill/SKILL.md) | v1.0.0 | 业务规则页 | 限制条件、办理要求 |
| [operation-steps-skill](operation-steps-skill/SKILL.md) | v1.0.0 | 操作步骤页 | 柜员/智能柜员机/Pad操作流程 |
| [key-points-skill](key-points-skill/SKILL.md) | v1.0.0 | 业务要点页 | 关键信息汇总 |
| [faq-skill](faq-skill/SKILL.md) | v1.0.0 | 常见问题页 | FAQ、异常处理 |

### 框架和工具

| 文件 | 功能 |
|------|------|
| [ppt-framework/styles.md](ppt-framework/styles.md) | 公共样式库：颜色/字体/间距/圆角、角色驱动字号、垂直节奏 |
| [ppt-framework/vertical-rhythm.md](ppt-framework/vertical-rhythm.md) | 五条带垂直节奏规范（上下文/主张/证据/含义/页脚） |
| [ppt-framework/vertical_rhythm.py](ppt-framework/vertical_rhythm.py) | 垂直节奏布局引擎，生成 5 种页面变体 |
| [ppt-framework/primitives/](ppt-framework/primitives/) | L3 原语组件库：text/panel/divider/badge/source_note/step_number |
| [ppt-framework/design-checklist.md](ppt-framework/design-checklist.md) | 9 项设计质量诊断清单（基于 jingmei-ppt 方法论） |
| [tools/design_validator.py](../tools/design_validator.py) | 设计质量自动化验证工具（9 项检查 + 面积重叠加权灰度重心） |
| [生成脚本/generate_page.py](../生成脚本/generate_page.py) | 端到端页面生成器：风格配方 + 垂直节奏 + 角色字号 → PPTX |

### 设计质量验证

```bash
# 运行 9 项设计诊断（自动输出检查报告）
py -3 tools/design_validator.py input.pptx

# 额外输出各页灰度重心（四象限权重 TL/TR/BL/BR）
py -3 tools/design_validator.py input.pptx --grayscale

# 导出 JSON 报告
py -3 tools/design_validator.py input.pptx --report report.json
```

检查项覆盖：标题主张、视觉停顿、视觉重心偏移、底部收尾、面板主次、图表解读、来源可读、缩略图层次、装饰功能。

## 使用流程

### 完整PPT生成流程

```
Word文档 → ①拆分章节 → ②逐章生成各类型页面 → ③合并为完整PPT
```

### 单页面生成流程

```
Word文档 → ①提取内容 → ②选择技能 → ③生成PPT页面
```

## 配色方案

### 预设配色

| 预设名称 | 主色 | 适用场景 |
|----------|------|----------|
| professional-blue | #1F3864 | 技术/业务类 |
| business-gray | #404040 | 管理/行政类 |
| vibrant-orange | #833C00 | 培训/宣导类 |

### 语义配色

| 颜色 | 色值 | 用途 |
|------|------|------|
| Success | #006100 | 成功、正常、通过 |
| Warning | #7F6000 | 警告、注意、判断 |
| Error | #C00000 | 错误、异常、禁止 |
| Info | #2E75B6 | 信息、提示、链接 |

## 设计规范

### 字体规范

| 元素 | 字号 | 字体 |
|------|------|------|
| H1 | 36pt | 微软雅黑 |
| H2 | 28pt | 微软雅黑 |
| H3 | 22pt | 微软雅黑 |
| H4 | 18pt | 微软雅黑 |
| Body | 14pt | 微软雅黑 |
| Caption | 12pt | 微软雅黑 |

### 间距规范

| 名称 | 值 |
|------|-----|
| xs | 4px |
| sm | 8px |
| md | 16px |
| lg | 24px |
| xl | 32px |
| xxl | 48px |

## 最佳实践

### 设计原则

1. **一致性**：保持颜色、字体、间距的一致性
2. **层次感**：通过字号、颜色、粗细区分信息层次
3. **留白**：适当留白，避免页面拥挤
4. **对齐**：保持元素对齐，提升可读性
5. **对比**：使用对比色突出重点信息

### 常见问题

1. **文字溢出**：控制文本长度，必要时截断或换行
2. **颜色不一致**：使用预设配色方案
3. **布局混乱**：遵循网格系统，保持对齐
4. **信息过载**：每页只传达一个核心信息

## 版本更新日志

### v1.1.0（2026-08-12）
- 引入 jingmei-ppt 方法论，执行 6 个敏捷迭代 Sprint
- 新增 style-brief-skill（风格简报 + 4 套配方：professional-blue/charcoal-minimal/warm-glow/red-alert）
- 新增 ppt-framework 公共样式库：垂直节奏五条带、角色驱动字号（claim/sectionLabel/body/annotation/source）
- 新增 vertical_rhythm.py 布局引擎与 primitives 原语组件库（6 个组件）
- 新增 design-checklist.md（9 项诊断）与 tools/design_validator.py（自动化验证 + 面积重叠加权灰度重心）
- 新增 生成脚本/generate_page.py：端到端生成器（风格配方 + 垂直节奏 + 角色字号 → PPTX）
- 闭环验证：用 generate_page.py 生成的「业务要点」页经 design_validator 评分 10.0/10（优秀）

### v1.0.0（2026-08-12）
- 首版发布：完整的PPT生成技能库
- 包含9个技能：封面、目录、表格、流程图、场景说明、业务规则、操作步骤、业务要点、常见问题
- 提供公共样式库和设计规范
- 支持3种配色方案
