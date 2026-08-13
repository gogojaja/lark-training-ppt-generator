#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
技能包结构标准化工具
为所有技能包添加标准目录结构
"""

import os
from pathlib import Path

# 技能包列表
SKILL_PACKAGES = [
    'cover-skill',
    'toc-skill',
    'scene-description-skill',
    'business-rules-skill',
    'operation-steps-skill',
    'key-points-skill',
    'faq-skill',
    'ppt-framework',
    'style-brief-skill',
    'table-skill',
    'document-processing',
    'flowchart-skill',
]

def create_standard_structure(skill_path):
    """为技能包创建标准结构"""
    created_dirs = []
    created_files = []

    # 创建 templates 目录
    templates_dir = skill_path / 'templates'
    if not templates_dir.exists():
        templates_dir.mkdir(parents=True, exist_ok=True)
        created_dirs.append('templates/')
        # 创建示例模板文件
        template_file = templates_dir / 'template_example.txt'
        template_file.write_text(f"""# {skill_path.name} 模板示例

## 功能说明
{skill_path.name} 模板

## 使用方法
1. 复制此模板
2. 根据需要修改内容
3. 使用技能包生成功能

## 示例
...
""", encoding='utf-8')
        created_files.append('templates/template_example.txt')

    # 创建 styles 目录
    styles_dir = skill_path / 'styles'
    if not styles_dir.exists():
        styles_dir.mkdir(parents=True, exist_ok=True)
        created_dirs.append('styles/')
        # 创建示例样式文件
        style_file = styles_dir / 'style_example.css'
        style_file.write_text(f"""/* {skill_path.name} 样式定义 */

/* 颜色定义 */
--primary-color: #1F3864;
--secondary-color: #2E75B6;
--accent-color: #ED7D31;

/* 字体定义 */
--font-family: 'Microsoft YaHei', sans-serif;
--font-size-base: 16px;

/* 间距定义 */
--spacing-xs: 4px;
--spacing-sm: 8px;
--spacing-md: 16px;
--spacing-lg: 24px;
--spacing-xl: 32px;
""", encoding='utf-8')
        created_files.append('styles/style_example.css')

    # 创建 examples 目录
    examples_dir = skill_path / 'examples'
    if not examples_dir.exists():
        examples_dir.mkdir(parents=True, exist_ok=True)
        created_dirs.append('examples/')
        # 创建示例文件
        example_file = examples_dir / 'example.py'
        example_file.write_text(f"""# {skill_path.name} 示例

"""
"""
# 示例 1: 基础使用
from {skill_path.name.replace('-', '_')} import SomeClass

obj = SomeClass()
result = obj.do_something()
print(result)

# 示例 2: 高级使用
from {skill_path.name.replace('-', '_')} import AdvancedClass

obj = AdvancedClass()
obj.configure(param1='value1')
obj.configure(param2='value2')
result = obj.process()
print(result)

# 示例 3: 错误处理
try:
    obj = SomeClass()
    result = obj.do_something()
except Exception as e:
    print(f'错误: {e}')
""", encoding='utf-8')
        created_files.append('examples/example.py')

    return created_dirs, created_files

def create_skill_index(skill_path):
    """创建 SKILL_INDEX.md"""
    skill_index_file = skill_path / 'SKILL_INDEX.md'
    if skill_index_file.exists():
        return False

    content = f"""# {skill_path.name} - 技能索引

## 元数据

- **技能名称**: {skill_path.name}
- **技能版本**: 1.0.0
- **发布日期**: 2026-08-13
- **技能描述**: {skill_path.name} 技能包
- **依赖技能**: 无

## 目录结构

- `templates/` - 模板文件目录
- `styles/` - 样式定义目录
- `examples/` - 示例文件目录

## 使用说明

### 基础使用

```python
from {skill_path.name.replace('-', '_')} import SkillClass

skill = SkillClass()
result = skill.execute()
```

### 高级使用

```python
from {skill_path.name.replace('-', '_')} import AdvancedSkill

skill = AdvancedSkill()
skill.configure(param1='value1')
result = skill.process()
```

## 示例

参见 `examples/` 目录中的示例文件。

## 依赖

无

## 更新日志

### v1.0.0 (2026-08-13)
- 初始版本
- 基础功能实现
"""

    skill_index_file.write_text(content, encoding='utf-8')
    return True

def create_readme(skill_path):
    """创建 README.md"""
    readme_file = skill_path / 'README.md'
    if readme_file.exists():
        return False

    content = f"""# {skill_path.name}

## 简介

{skill_path.name} 是一个用于 {skill_path.name.replace('-', ' ')} 的技能包。

## 功能列表

- 功能 1
- 功能 2
- 功能 3

## 使用方法

### 安装

```bash
# 克隆仓库
git clone <repository-url>

# 安装依赖
pip install -r requirements.txt
```

### 基础使用

```python
from {skill_path.name.replace('-', '_')} import SkillClass

skill = SkillClass()
result = skill.execute()
print(result)
```

### 高级使用

```python
from {skill_path.name.replace('-', '_')} import AdvancedSkill

skill = AdvancedSkill()
skill.configure(param1='value1')
result = skill.process()
print(result)
```

## 示例

参见 `examples/` 目录中的示例文件。

## 配置

### 配置文件

```json
{{
  "param1": "value1",
  "param2": "value2"
}}
```

### 环境变量

- `SKILL_PARAM1`: 参数1的值
- `SKILL_PARAM2`: 参数2的值

## 依赖

- Python >= 3.8
- 其他依赖...

## 许可证

MIT License

## 贡献

欢迎贡献！请提交 Pull Request。
"""

    readme_file.write_text(content, encoding='utf-8')
    return True

def main():
    """主函数"""
    print("=" * 60)
    print("技能包结构标准化工具")
    print("=" * 60)

    total_created_dirs = 0
    total_created_files = 0
    total_updated_files = 0

    for skill_name in SKILL_PACKAGES:
        skill_path = Path('skills') / skill_name
        if not skill_path.exists():
            print(f"\n技能包不存在: {skill_name}")
            continue

        print(f"\n处理技能包: {skill_name}")

        # 创建标准结构
        created_dirs, created_files = create_standard_structure(skill_path)
        if created_dirs:
            print(f"  [+] 创建目录: {', '.join(created_dirs)}")
            total_created_dirs += len(created_dirs)

        if created_files:
            print(f"  [+] 创建文件: {', '.join(created_files)}")
            total_created_files += len(created_files)

        # 创建 SKILL_INDEX.md
        if create_skill_index(skill_path):
            print(f"  [+] 创建 SKILL_INDEX.md")
            total_updated_files += 1

        # 创建 README.md
        if create_readme(skill_path):
            print(f"  [+] 创建 README.md")
            total_updated_files += 1

    print("\n" + "=" * 60)
    print("标准化完成")
    print("=" * 60)
    print(f"创建目录数: {total_created_dirs}")
    print(f"创建文件数: {total_created_files}")
    print(f"更新文件数: {total_updated_files}")

if __name__ == '__main__':
    main()
