# Skill: cover-skill

# 封面页生成技能

## 1. 元数据

- **技能名称**：cover-skill
- **技能版本**：v1.0.0
- **发布日期**：2026-08-12
- **技能定位**：生成专业的PPT封面页，包含标题、副标题、日期等信息
- **适用场景**：培训PPT、汇报PPT、方案PPT的封面

## 2. 设计规范

### 2.1 布局模板

**模板A：标准封面**
```
┌─────────────────────────────────────┐
│ ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ │ ← 顶部装饰条
│                                     │
│                                     │
│         [主标题]                     │ ← 36pt，加粗，白色
│         ─────────                   │ ← 分隔线
│         [副标题]                     │ ← 18pt，浅蓝色
│                                     │
│                                     │
│         [日期/版本/作者]             │ ← 12pt，浅灰色
│                                     │
└─────────────────────────────────────┘
```

**模板B：左侧装饰封面**
```
┌─────────────────────────────────────┐
│ ████████                            │ ← 左侧装饰块
│ ████████                            │
│ ████████    [主标题]                 │
│ ████████    [副标题]                 │
│ ████████                            │
│ ████████    [日期/版本]              │
│                                     │
└─────────────────────────────────────┘
```

**模板C：简约封面**
```
┌─────────────────────────────────────┐
│                                     │
│                                     │
│                                     │
│         [主标题]                     │
│         [副标题]                     │
│                                     │
│                                     │
│ ─────────────────────────────────── │ ← 底部信息栏
│ [日期]  [版本]  [作者]               │
└─────────────────────────────────────┘
```

### 2.2 配色方案

**专业蓝（默认）**：
- 背景色：#1F3864（深蓝）
- 主标题：#FFFFFF（白色）
- 副标题：#93C5FD（浅蓝）
- 装饰条：#2E75B6（中蓝）

**商务灰**：
- 背景色：#404040（深灰）
- 主标题：#FFFFFF（白色）
- 副标题：#D9D9D9（浅灰）
- 装饰条：#808080（中灰）

**活力橙**：
- 背景色：#833C00（深橙）
- 主标题：#FFFFFF（白色）
- 副标题：#FCE4D6（浅橙）
- 装饰条：#ED7D31（中橙）

### 2.3 字体规范

| 元素 | 字号 | 字体 | 颜色 | 对齐 |
|------|------|------|------|------|
| 主标题 | 36pt | 微软雅黑 | 白色 | 居中 |
| 副标题 | 18pt | 微软雅黑 | 浅蓝 | 居中 |
| 日期 | 12pt | 微软雅黑 | 浅灰 | 居中 |
| 版本 | 12pt | 微软雅黑 | 浅灰 | 右对齐 |
| 作者 | 12pt | 微软雅黑 | 浅灰 | 左对齐 |

## 3. 工作流

```
输入信息 → ①选择模板 → ②配置样式 → ③生成PPT
```

### Step 1：收集封面信息

**必需信息**：
- 主标题：PPT的主要标题
- 副标题：补充说明（可选）

**可选信息**：
- 日期：制作日期
- 版本：版本号
- 作者：制作人/部门

### Step 2：选择模板和配色

**模板选择原则**：
- 标准封面：通用场景
- 左侧装饰：正式汇报
- 简约封面：简洁风格

**配色选择原则**：
- 专业蓝：技术/业务类
- 商务灰：管理/行政类
- 活力橙：培训/宣导类

### Step 3：生成封面页

**生成脚本**：
```bash
# 使用默认模板
py -3 生成脚本/gen_cover.py --title "个人综合签约业务培训" --out 封面.pptx

# 指定模板和配色
py -3 生成脚本/gen_cover.py --title "业务培训" --subtitle "操作手册" --template standard --preset blue --out 封面.pptx

# 完整参数
py -3 生成脚本/gen_cover.py --title "业务培训" --subtitle "操作手册" --date "2026-08-12" --version "v1.0" --author "培训部" --template sidebar --preset gray --out 封面.pptx
```

## 4. 代码实现

### 4.1 XML结构

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
      </p:bgRef>
    </p:bg>
    <p:spTree>
      <!-- 装饰条 -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="deco_bar"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="166688"/></a:xfrm>
          <a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill>
        </p:spPr>
      </p:sp>
      
      <!-- 主标题 -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="title"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1000000" y="2500000"/><a:ext cx="10192000" cy="1500000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" anchor="ctr"/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="zh-CN" sz="3600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              </a:rPr>
              <a:t>主标题</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      
      <!-- 副标题 -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="subtitle"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1000000" y="4000000"/><a:ext cx="10192000" cy="800000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" anchor="ctr"/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="zh-CN" sz="1800" dirty="0">
                <a:solidFill><a:srgbClr val="93C5FD"/></a:solidFill>
              </a:rPr>
              <a:t>副标题</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
```

## 5. 最佳实践

### 5.1 设计原则

1. **简洁明了**：封面信息不宜过多，突出核心标题
2. **视觉层次**：主标题最醒目，副标题次之，辅助信息最小
3. **颜色对比**：确保文字与背景有足够对比度
4. **留白适当**：避免页面拥挤，保持呼吸感

### 5.2 常见问题

1. **文字溢出**：控制标题长度，必要时换行
2. **颜色不协调**：使用预设配色方案
3. **布局不美观**：遵循网格对齐原则
4. **信息不完整**：确保包含必要信息

### 5.3 优化建议

1. **添加Logo**：在封面添加公司/部门Logo
2. **背景图片**：使用合适的背景图片增强视觉效果
3. **动画效果**：添加简单的进入动画
4. **响应式设计**：考虑不同屏幕尺寸的显示效果

## 6. 版本更新日志

### v1.0.0（2026-08-12）
- 首版发布：支持3种封面模板
- 支持3种配色方案
- 支持主标题、副标题、日期、版本、作者信息
