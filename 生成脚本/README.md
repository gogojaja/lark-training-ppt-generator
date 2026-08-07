# PPT生成器 v5.0.0

配置驱动、组件化架构的培训PPT生成器

## 快速开始

### 1. 安装依赖

```bash
cd 生成脚本
npm install
```

### 2. 配置

编辑 `config.yaml`:

```yaml
theme: ocean-gradient
export_format: pptx
page_size: "16:9"
output_dir: "生成产物"
```

### 3. 生成PPT

```bash
# 使用默认配置
node generate.js

# 指定主题
node generate.js --theme charcoal-minimal

# 指定配置文件
node generate.js --config config.yaml

# 指定输出路径
node generate.js --output my_presentation.pptx
```

## 目录结构

```
生成脚本/
├── components/           # 可复用组件
│   ├── ThemeManager.js   # 主题管理器
│   ├── ComponentLoader.js # 组件加载器
│   ├── SchemaValidator.js # Schema验证器
│   ├── ConfigManager.js  # 配置管理器
│   ├── SlideRenderer.js  # 幻灯片渲染器
│   ├── Orchestrator.js   # 主控制器
│   ├── CoverComponent.js # 封面组件
│   └── InfoCardComponent.js # 信息卡片组件
├── themes/               # 主题配置
│   ├── ocean-gradient.json
│   └── charcoal-minimal.json
├── schemas/              # JSON Schema
│   ├── slide_plan.json
│   ├── theme.json
│   └── config.json
├── config.yaml           # 用户配置
├── slide_plan.json       # 幻灯片规划
├── generate.js           # 主生成器入口
├── validate.js           # 验证脚本
├── test.js               # 测试脚本
└── package.json          # 项目配置
```

## 自定义主题

### 创建新主题

在 `themes/` 目录创建JSON文件:

```json
{
  "name": "My Theme",
  "palette": {
    "primary": "0052D9",
    "secondary": "00B96B",
    "accent": "FF7D00",
    "bg": "F5F7FA",
    "text": "1F2937"
  },
  "fonts": {
    "heading": "Microsoft YaHei",
    "body": "Microsoft YaHei"
  }
}
```

### 使用主题

```bash
node generate.js --theme my-theme
```

## 验证配置

```bash
# 验证所有配置文件
node validate.js

# 运行测试
node test.js
```

## 组件开发

### 添加新组件

1. 在 `components/` 目录创建新文件
2. 实现 `render(slide, theme)` 方法
3. 在 `SlideRenderer.js` 中添加类型映射

### 组件接口

```javascript
class MyComponent {
    render(slide, theme) {
        return {
            type: 'my_type',
            elements: [...],
            notes: '...'
        };
    }
}
```

## 迁移指南

### 从v4.x迁移到v5.x

1. **配置外部化**: 颜色/字体配置移到 `themes/*.json`
2. **组件化**: `generate.js` 拆分为独立组件
3. **Schema验证**: 新增JSON Schema验证

### 向后兼容

- 现有 `slide_plan.json` 格式保持兼容
- 现有命令行参数保持兼容
- 新功能通过配置启用

## 故障排除

### 常见问题

1. **主题文件不存在**: 检查 `themes/` 目录
2. **组件加载失败**: 检查 `components/` 目录
3. **Schema验证失败**: 检查JSON格式

### 调试模式

```bash
DEBUG=1 node generate.js
```

## 更新日志

### v5.0.0 (2026-08-07)
- 新增配置驱动主题系统
- 新增组件化架构
- 新增Schema验证
- 新增多方案支持
- 保持向后兼容
