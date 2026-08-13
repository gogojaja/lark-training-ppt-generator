#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档整理脚本
自动整理项目文档到正确的位置
"""

import os
import shutil
import sys
from pathlib import Path

# 定义源目录和目标目录
SOURCE_DIRS = {
    '生成脚本': '生成脚本',
    '工具': 'tools',
    '输入文档': '输入文档',
    '生成产物/表格': '生成产物/表格',
    '生成产物/演示': '生成产物/演示',
    '生成产物/输出': '输出文档',
}

# 技能包目录
SKILL_DIRS = [
    'skills/cover-skill',
    'skills/toc-skill',
    'skills/scene-description-skill',
    'skills/business-rules-skill',
    'skills/operation-steps-skill',
    'skills/key-points-skill',
    'skills/faq-skill',
    'skills/ppt-framework',
    'skills/style-brief-skill',
    'skills/table-skill',
    'skills/document-processing',
    'skills/flowchart-skill',
]

def check_directory_structure():
    """检查目录结构"""
    print("=" * 60)
    print("检查目录结构")
    print("=" * 60)

    required_dirs = [
        '输入文档',
        '生成脚本',
        '生成产物',
        '生成产物/表格',
        '生成产物/演示',
        '生成产物/输出',
        '输出文档',
        '技能包',
        '工具',
        '文档',
    ]

    missing_dirs = []
    for dir_name in required_dirs:
        dir_path = Path(dir_name)
        if not dir_path.exists():
            missing_dirs.append(dir_name)

    if missing_dirs:
        print("\n缺少的目录：")
        for dir_name in missing_dirs:
            print(f"  - {dir_name}")
        print("\n正在创建缺失的目录...")
        for dir_name in missing_dirs:
            Path(dir_name).mkdir(parents=True, exist_ok=True)
        print("目录创建完成！")
    else:
        print("\n目录结构完整！")

def organize_generated_files():
    """整理生成的文件"""
    print("\n" + "=" * 60)
    print("整理生成的文件")
    print("=" * 60)

    # 整理根目录的 PPT 文件
    root_files = list(Path('.').glob('*.pptx'))
    if root_files:
        print(f"\n发现根目录下的 PPT 文件：")
        for file in root_files:
            print(f"  - {file.name}")
            # 移动到生成产物/输出/
            dest_dir = Path('生成产物/输出')
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest_file = dest_dir / file.name
            shutil.move(str(file), str(dest_file))
            print(f"    -> {dest_file}")

    # 整理根目录的 JSON 文件（表格数据）
    json_files = list(Path('.').glob('*.json'))
    if json_files:
        print(f"\n发现根目录下的 JSON 文件：")
        for file in json_files:
            print(f"  - {file.name}")
            # 检查是否是表格数据
            if '表格' in file.name or 'data' in file.name.lower():
                dest_dir = Path('生成产物/表格')
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest_file = dest_dir / file.name
                shutil.move(str(file), str(dest_file))
                print(f"    -> {dest_file}")
            else:
                print(f"    -> 保留在根目录")

def organize_temp_files():
    """整理临时文件"""
    print("\n" + "=" * 60)
    print("整理临时文件")
    print("=" * 60)

    # 删除 __pycache__ 目录
    pycache_dirs = list(Path('.').rglob('__pycache__'))
    if pycache_dirs:
        print(f"\n发现 __pycache__ 目录：")
        for dir_path in pycache_dirs:
            print(f"  - {dir_path}")
            shutil.rmtree(dir_path)
            print(f"    -> 已删除")

    # 删除 .bak 文件
    bak_files = list(Path('.').rglob('*.bak*'))
    if bak_files:
        print(f"\n发现 .bak 备份文件：")
        for file in bak_files:
            print(f"  - {file.name}")
            file.unlink()
            print(f"    -> 已删除")

def show_summary():
    """显示整理摘要"""
    print("\n" + "=" * 60)
    print("整理摘要")
    print("=" * 60)

    print("\n目录结构：")
    print("  输入文档/          - 原始输入文档")
    print("  生成脚本/          - 生成脚本")
    print("  生成产物/          - 生成产物")
    print("    ├── 表格/        - 表格相关产物")
    print("    ├── 演示/        - 演示文稿")
    print("    └── 输出/        - 最终输出")
    print("  输出文档/          - 最终输出文档")
    print("  技能包/            - 技能包目录")
    print("  工具/              - 辅助工具")
    print("  文档/              - 项目文档")

    print("\n文件命名规范：")
    print("  - Python 脚本：小写字母 + 下划线（generate_page.py）")
    print("  - JSON 数据：清晰描述（账户修改功能适用范围.json）")
    print("  - PPT 文件：清晰描述（账户修改功能适用范围.pptx）")
    print("  - 文档文件：清晰描述（使用手册.md）")

def main():
    """主函数"""
    print("=" * 60)
    print("文档整理工具")
    print("=" * 60)

    try:
        # 1. 检查目录结构
        check_directory_structure()

        # 2. 整理生成的文件
        organize_generated_files()

        # 3. 整理临时文件
        organize_temp_files()

        # 4. 显示摘要
        show_summary()

        print("\n" + "=" * 60)
        print("整理完成！")
        print("=" * 60)

    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
