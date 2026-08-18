---
name: "ppt-framework"
description: "PPT框架技能包：页面生成、样式管理、布局系统、垂直节奏。触发词：页面生成、样式管理、布局系统、垂直节奏。Load when the user wants to generate PPT pages, manage styles, or implement layout systems."
---

# PPTFrameworkSkill PPT框架技能包

> 版权：`../references/COPYRIGHT.md`　Token：`../references/token_standard.md`

## 1. 元数据

- **技能名称**：ppt-framework
- **技能版本**：v1.0.0
- **发布日期**：2026-08-13
- **参考标准**：Jingmei-PPT Methodology

## 2. 触发规则

用户表达「页面生成/样式管理/布局系统/垂直节奏」时加载本包。

## 3. 流程（路由到 domain/）

| 环节 | action | 明细 |
|------|--------|------|
| 页面生成 | generate_page | `domain/page_generator.py` |
| 样式管理 | manage_style | `domain/style_manager.py` |
| 布局系统 | layout_system | `domain/layout.py` |
| 垂直节奏 | vertical_rhythm | `domain/vertical_rhythm.py` |

## 4. 核心能力

### 4.1 页面生成

使用 `generate_page.py` 生成PPT页面：

```bash
# 生成单页PPT
py -3 生成脚本/generate_page.py --page-type cover --input input.json --output output.pptx

# 生成多页PPT
py -3 生成脚本/generate_page.py --page-type toc --input input.json --output output.pptx
```

### 4.2 样式管理

使用 `styles.md` 定义样式：

- 定义颜色、字体、间距
- 支持主题切换
- 支持自定义样式

### 4.3 布局系统

使用 `primitives/` 中的布局组件：

- 基础布局：Flex、Grid
- 垂直布局：垂直堆叠、水平排列
- 响应式布局

### 4.4 垂直节奏

使用 `vertical_rhythm.py` 实现垂直节奏：

```bash
# 设置垂直节奏
py -3 生成脚本/vertical_rhythm.py --base 16 --ratio 1.5
```

## 5. 技术实现

### 5.1 页面生成器

使用 pptxgenjs 库：

- 支持多种页面类型
- 支持自定义样式
- 支持模板

### 5.2 样式系统

基于垂直节奏系统：

```python
# 基础尺寸
base_size = 16  # 基础字号
line_height = base_size * 1.5  # 行高

# 垂直间距
spacing_unit = base_size  # 基础间距单位
```

### 5.3 布局系统

使用 CSS Flexbox 和 Grid：

- Flexbox：用于一维布局
- Grid：用于二维布局
- 响应式：支持不同屏幕尺寸

## 6. 输出规范

- PPTX文件：标准PowerPoint格式
- 样式文件：CSS/JSON格式
- 布局文件：JSON格式

## 7. 边界

- 仅支持PPTX格式
- 需要安装pptxgenjs库
- 复杂动画需要手动调整

---

**文档版本**：v1.0.0　**最后更新**：2026-08-13
**知识产权所有**：段波（验证邮箱：duanbo.douglas@163.com）
