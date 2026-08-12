# PPT生成框架 - 公共样式库

## 1. 设计规范

### 1.1 颜色系统

**主色调**：
| 颜色名称 | 色值 | 用途 |
|----------|------|------|
| Primary Dark | #1F3864 | 标题栏、强调色 |
| Primary Medium | #2E75B6 | 次要强调、边框 |
| Primary Light | #D6E4F0 | 浅色背景、装饰 |

**语义色**：
| 颜色名称 | 色值 | 用途 |
|----------|------|------|
| Success | #006100 | 成功、正常、通过 |
| Warning | #7F6000 | 警告、注意、判断 |
| Error | #C00000 | 错误、异常、禁止 |
| Info | #2E75B6 | 信息、提示、链接 |

**中性色**：
| 颜色名称 | 色值 | 用途 |
|----------|------|------|
| Text Primary | #333333 | 正文文字 |
| Text Secondary | #666666 | 次要文字 |
| Text Muted | #999999 | 辅助文字 |
| Background | #FFFFFF | 页面背景 |
| Surface | #F8FAFC | 卡片背景 |
| Border | #E0E0E0 | 边框、分隔线 |

### 1.2 字体系统

**字体家族**：
- 标题：Microsoft YaHei (微软雅黑)
- 正文：Microsoft YaHei (微软雅黑)
- 数字/代码：Consolas

**角色驱动的字号系统**（基于jingmei-ppt方法论）：

> 核心理念：组件设计时思考的是"这个文本承担什么角色"，而不是"这个文本应该用H几"

| 角色 | 字号 | 字重 | 行高 | 用途 |
|------|------|------|------|------|
| claim | 32pt | bold | 1.2 | 页面核心结论/主张 |
| sectionLabel | 12pt | normal | 1.3 | 章节标签、面包屑、页码 |
| body | 14pt | normal | 1.5 | 正文内容 |
| annotation | 12pt | normal | 1.4 | 注释、说明 |
| source | 10pt | normal | 1.3 | 来源标注 |

**传统字号映射**（向后兼容）：
| 元素 | 字号 | 行高 | 用途 |
|------|------|------|------|
| H1 | 36pt | 1.2 | 页面主标题 |
| H2 | 28pt | 1.3 | 章节标题 |
| H3 | 22pt | 1.4 | 小节标题 |
| H4 | 18pt | 1.4 | 卡片标题 |
| Body | 14pt | 1.5 | 正文内容 |
| Caption | 12pt | 1.4 | 说明文字 |
| Small | 10pt | 1.3 | 辅助信息 |

### 1.3 间距系统

**间距规范**：
| 名称 | 值 | 用途 |
|------|-----|------|
| xs | 4px | 极小间距 |
| sm | 8px | 小间距 |
| md | 16px | 中等间距 |
| lg | 24px | 大间距 |
| xl | 32px | 特大间距 |
| xxl | 48px | 超大间距 |

### 1.4 圆角系统

**圆角规范**：
| 名称 | 值 | 用途 |
|------|-----|------|
| none | 0 | 无圆角 |
| sm | 4px | 小圆角 |
| md | 8px | 中等圆角 |
| lg | 12px | 大圆角 |
| full | 9999px | 全圆角（圆形） |

### 1.5 垂直节奏系统

> 基于 jingmei-ppt 方法论，详见 [vertical-rhythm.md](vertical-rhythm.md)

**五条带结构**：
```
┌─────────────────────────────────────┐ ← 0%
│  ① 上下文条带 (Context Band)        │ ← 0-8%
├─────────────────────────────────────┤ ← 8%
│  ② 主张条带 (Claim Band)           │ ← 8-15%
├─────────────────────────────────────┤ ← 15%
│  ③ 证据区 (Evidence Zone)          │ ← 15-82%
├─────────────────────────────────────┤ ← 82%
│  ④ 含义/来源条带 (Meaning Band)    │ ← 82-92%
├─────────────────────────────────────┤ ← 92%
│  ⑤ 页脚安全区 (Footer Safe Zone)   │ ← 92-100%
└─────────────────────────────────────┘ ← 100%
```

**条带功能**：
| 条带 | 位置 | 功能 | 内容 |
|------|------|------|------|
| 上下文条带 | 0-8% | 页面定位 | 章节标签、页码、Logo |
| 主张条带 | 8-15% | 核心结论 | 页面主标题、副标题 |
| 证据区 | 15-82% | 支撑内容 | 图表、表格、文本、卡片 |
| 含义条带 | 82-92% | 解读行动 | 数据解读、行动建议、来源 |
| 页脚安全区 | 92-100% | 收尾 | 页码、Logo、版权 |

**设计规则**：
1. 条带顺序固定：上下文 → 主张 → 证据 → 含义 → 页脚
2. 主张条带必须有内容
3. 底部必须收尾
4. 证据区不能溢出

## 2. 组件库

### 2.1 标题栏组件

**样式A：深色背景**
```xml
<p:sp>
  <p:spPr>
    <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:p>
      <a:r>
        <a:rPr sz="2800" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        </a:rPr>
        <a:t>页面标题</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

**样式B：浅色背景**
```xml
<p:sp>
  <p:spPr>
    <a:solidFill><a:srgbClr val="D6E4F0"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:p>
      <a:r>
        <a:rPr sz="2800" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
        </a:rPr>
        <a:t>页面标题</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

### 2.2 卡片组件

**基础卡片**：
```xml
<p:sp>
  <p:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91440" rIns="91440" tIns="91440" bIns="91440"/>
  </p:txBody>
</p:sp>
```

**带阴影卡片**：
```xml
<p:sp>
  <p:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>
    <a:effectLst>
      <a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0">
        <a:srgbClr val="000000"><a:alpha val="23000"/></a:srgbClr>
      </a:outerShdw>
    </a:effectLst>
  </p:spPr>
</p:sp>
```

### 2.3 列表组件

**无序列表**：
```xml
<a:p>
  <a:pPr marL="228600" indent="-228600"/>
  <a:r>
    <a:rPr sz="1400" dirty="0">
      <a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill>
    </a:rPr>
    <a:t>• </a:t>
  </a:r>
  <a:r>
    <a:rPr sz="1400" dirty="0"/>
    <a:t>列表项内容</a:t>
  </a:r>
</a:p>
```

**有序列表**：
```xml
<a:p>
  <a:pPr marL="228600" indent="-228600"/>
  <a:r>
    <a:rPr sz="1400" b="1" dirty="0">
      <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
    </a:rPr>
    <a:t>1. </a:t>
  </a:r>
  <a:r>
    <a:rPr sz="1400" dirty="0"/>
    <a:t>列表项内容</a:t>
  </a:r>
</a:p>
```

### 2.4 表格组件

**基础表格**：
```xml
<p:tbl>
  <p:tblPr>
    <a:tblW w="9144000" type="dist"/>
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
    <p:gridCol w="3048000"/>
    <p:gridCol w="3048000"/>
    <p:gridCol w="3048000"/>
  </p:tblGrid>
  <p:tr>
    <p:trPr><a:trHeight val="457200"/></p:trPr>
    <p:tc>
      <p:tcPr>
        <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
      </p:tcPr>
      <p:txBody>
        <a:p>
          <a:r>
            <a:rPr sz="1400" b="1" dirty="0">
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            </a:rPr>
            <a:t>表头</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:tc>
  </p:tr>
</p:tbl>
```

## 3. 页面模板

### 3.1 封面页

**布局**：
```
┌─────────────────────────────────────┐
│                                     │
│                                     │
│         [主标题]                     │
│         [副标题]                     │
│                                     │
│         [日期/版本]                  │
│                                     │
└─────────────────────────────────────┘
```

**规范**：
- 主标题：36pt，加粗，白色
- 副标题：18pt，浅蓝色
- 背景：深蓝色 #1F3864

### 3.2 目录页

**布局**：
```
┌─────────────────────────────────────┐
│  目录                               │
├─────────────────────────────────────┤
│  01  章节标题一                      │
│  02  章节标题二                      │
│  03  章节标题三                      │
│  04  章节标题四                      │
└─────────────────────────────────────┘
```

**规范**：
- 章节编号：24pt，加粗，主色
- 章节标题：18pt，深灰色

### 3.3 内容页

**布局**：
```
┌─────────────────────────────────────┐
│  页面标题                            │
├─────────────────────────────────────┤
│                                     │
│  [内容区域]                         │
│                                     │
└─────────────────────────────────────┘
```

**规范**：
- 标题栏：高度 80px，深色背景
- 内容区域： padding 24px

### 3.4 分隔页

**布局**：
```
┌─────────────────────────────────────┐
│                                     │
│         PART 01                     │
│         章节标题                     │
│                                     │
└─────────────────────────────────────┘
```

**规范**：
- PART编号：48pt，浅色
- 章节标题：36pt，深色
- 背景：白色或浅色

## 4. 工具函数

### 4.1 颜色转换

```python
def hex_to_rgb(hex_color):
    """将16进制颜色转换为RGB"""
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(r, g, b):
    """将RGB转换为16进制颜色"""
    return f"{r:02x}{g:02x}{b:02x}"

def lighten_color(hex_color, factor=0.2):
    """提亮颜色"""
    r, g, b = hex_to_rgb(hex_color)
    r = min(255, int(r + (255 - r) * factor))
    g = min(255, int(g + (255 - g) * factor))
    b = min(255, int(b + (255 - b) * factor))
    return rgb_to_hex(r, g, b)

def darken_color(hex_color, factor=0.2):
    """加深颜色"""
    r, g, b = hex_to_rgb(hex_color)
    r = max(0, int(r * (1 - factor)))
    g = max(0, int(g * (1 - factor)))
    b = max(0, int(b * (1 - factor)))
    return rgb_to_hex(r, g, b)
```

### 4.2 尺寸转换

```python
def cm_to_emu(cm):
    """将厘米转换为EMU"""
    return int(cm * 360000)

def emu_to_cm(emu):
    """将EMU转换为厘米"""
    return emu / 360000

def pt_to_emu(pt):
    """将磅转换为EMU"""
    return int(pt * 12700)
```

### 4.3 文本处理

```python
def truncate_text(text, max_length=15):
    """截断文本"""
    if len(text) <= max_length:
        return text
    return text[:max_length-2] + "..."

def wrap_text(text, max_width=10):
    """文本换行"""
    lines = []
    current_line = ""
    for char in text:
        current_line += char
        if len(current_line) >= max_width:
            lines.append(current_line)
            current_line = ""
    if current_line:
        lines.append(current_line)
    return "\n".join(lines)
```

## 5. 最佳实践

### 5.1 设计原则

1. **一致性**：保持颜色、字体、间距的一致性
2. **层次感**：通过字号、颜色、粗细区分信息层次
3. **留白**：适当留白，避免页面拥挤
4. **对齐**：保持元素对齐，提升可读性
5. **对比**：使用对比色突出重点信息

### 5.2 常见问题

1. **文字溢出**：控制文本长度，必要时截断或换行
2. **颜色不一致**：使用预设配色方案
3. **布局混乱**：遵循网格系统，保持对齐
4. **信息过载**：每页只传达一个核心信息

### 5.3 优化建议

1. **使用图标**：适当使用图标增强视觉效果
2. **动画效果**：添加简单的进入动画
3. **响应式设计**：考虑不同屏幕尺寸的显示效果
4. **无障碍设计**：确保颜色对比度符合标准
