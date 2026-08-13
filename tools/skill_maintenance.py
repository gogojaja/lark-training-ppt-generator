#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
技能维护工具
评估和改进技能包质量
"""

import os
import json
from pathlib import Path
from datetime import datetime

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

def check_skill_structure(skill_path):
    """检查技能包结构"""
    issues = []

    # 检查 SKILL.md 是否存在
    skill_md = skill_path / 'SKILL.md'
    if not skill_md.exists():
        issues.append('缺少 SKILL.md 文件')

    # 检查是否有模板目录
    templates_dir = skill_path / 'templates'
    if not templates_dir.exists():
        issues.append('缺少 templates 目录')

    # 检查是否有样式目录
    styles_dir = skill_path / 'styles'
    if not styles_dir.exists():
        issues.append('缺少 styles 目录')

    # 检查 SKILL_INDEX.md 是否存在
    skill_index = skill_path / 'SKILL_INDEX.md'
    if not skill_index.exists():
        issues.append('缺少 SKILL_INDEX.md 文件')

    # 检查是否有示例文件
    examples_dir = skill_path / 'examples'
    if not examples_dir.exists():
        issues.append('缺少 examples 目录')

    # 检查是否有 README.md
    readme = skill_path / 'README.md'
    if not readme.exists():
        issues.append('缺少 README.md 文件')

    return issues

def check_skill_content(skill_path):
    """检查技能包内容"""
    issues = []

    # 读取 SKILL.md
    skill_md = skill_path / 'SKILL.md'
    if skill_md.exists():
        with open(skill_md, 'r', encoding='utf-8') as f:
            content = f.read()
            # 检查是否包含基本元数据
            if '技能名称' not in content and 'Skill Name' not in content:
                issues.append('SKILL.md 缺少技能名称')
            if '技能版本' not in content and 'Version' not in content:
                issues.append('SKILL.md 缺少版本信息')
            if '发布日期' not in content and 'Date' not in content:
                issues.append('SKILL.md 缺少发布日期')

    return issues

def check_skill_files(skill_path):
    """检查技能包文件"""
    issues = []

    # 列出所有文件
    files = list(skill_path.rglob('*'))

    # 检查是否有临时文件
    for file in files:
        if file.is_file():
            if file.name.endswith('.pyc') or file.name.endswith('__pycache__'):
                issues.append(f'发现临时文件: {file.relative_to(skill_path)}')
            if file.name.startswith('~$'):
                issues.append(f'发现临时文件: {file.relative_to(skill_path)}')
            if file.name.endswith('.bak'):
                issues.append(f'发现备份文件: {file.relative_to(skill_path)}')

    return issues

def check_skill_completeness(skill_path):
    """检查技能包完整性"""
    issues = []

    # 检查必需的文件
    required_files = [
        'SKILL.md',
        'templates/',
        'styles/',
        'SKILL_INDEX.md',
    ]

    for required in required_files:
        file_path = skill_path / required
        if not file_path.exists():
            issues.append(f'缺少必需文件: {required}')

    return issues

def analyze_skill(skill_name, skill_path):
    """分析单个技能包"""
    print(f"\n{'=' * 60}")
    print(f"分析技能: {skill_name}")
    print(f"{'=' * 60}")

    # 检查结构
    print("\n1. 检查结构...")
    structure_issues = check_skill_structure(skill_path)
    if structure_issues:
        print(f"   [!] 发现 {len(structure_issues)} 个结构问题:")
        for issue in structure_issues:
            print(f"      - {issue}")
    else:
        print("   [OK] 结构完整")

    # 检查内容
    print("\n2. 检查内容...")
    content_issues = check_skill_content(skill_path)
    if content_issues:
        print(f"   [!] 发现 {len(content_issues)} 个内容问题:")
        for issue in content_issues:
            print(f"      - {issue}")
    else:
        print("   [OK] 内容完整")

    # 检查文件
    print("\n3. 检查文件...")
    file_issues = check_skill_files(skill_path)
    if file_issues:
        print(f"   [!] 发现 {len(file_issues)} 个文件问题:")
        for issue in file_issues:
            print(f"      - {issue}")
    else:
        print("   [OK] 文件整洁")

    # 检查完整性
    print("\n4. 检查完整性...")
    completeness_issues = check_skill_completeness(skill_path)
    if completeness_issues:
        print(f"   [!] 发现 {len(completeness_issues)} 个完整性问题:")
        for issue in completeness_issues:
            print(f"      - {issue}")
    else:
        print("   [OK] 完整性良好")

    # 统计文件数量
    files = list(skill_path.rglob('*'))
    files = [f for f in files if f.is_file()]
    print(f"\n   文件统计:")
    print(f"      总文件数: {len(files)}")
    print(f"      目录数: {len([f for f in files if f.is_dir()])}")
    print(f"      文件数: {len([f for f in files if f.is_file()])}")

    return {
        'name': skill_name,
        'structure_issues': len(structure_issues),
        'content_issues': len(content_issues),
        'file_issues': len(file_issues),
        'completeness_issues': len(completeness_issues),
        'total_issues': len(structure_issues) + len(content_issues) + len(file_issues) + len(completeness_issues),
        'total_files': len(files),
    }

def generate_report(analysis_results):
    """生成分析报告"""
    print("\n" + "=" * 60)
    print("技能包分析报告")
    print("=" * 60)

    # 按问题数量排序
    sorted_results = sorted(analysis_results, key=lambda x: x['total_issues'])

    print(f"\n技能包统计:")
    print(f"  总技能包数量: {len(analysis_results)}")
    print(f"  无问题的技能包: {len([r for r in analysis_results if r['total_issues'] == 0])}")
    print(f"  有问题的技能包: {len([r for r in analysis_results if r['total_issues'] > 0])}")

    print(f"\n问题统计:")
    print(f"  结构问题总数: {sum(r['structure_issues'] for r in analysis_results)}")
    print(f"  内容问题总数: {sum(r['content_issues'] for r in analysis_results)}")
    print(f"  文件问题总数: {sum(r['file_issues'] for r in analysis_results)}")
    print(f"  完整性问题总数: {sum(r['completeness_issues'] for r in analysis_results)}")
    print(f"  总问题数: {sum(r['total_issues'] for r in analysis_results)}")

    print(f"\n问题最多的技能包 (前 5 名):")
    for i, result in enumerate(sorted_results[:5], 1):
        if result['total_issues'] > 0:
            print(f"  {i}. {result['name']}: {result['total_issues']} 个问题")

    # 详细报告
    print(f"\n详细报告:")
    for result in sorted_results:
        if result['total_issues'] > 0:
            print(f"\n  {result['name']}:")
            if result['structure_issues'] > 0:
                print(f"    结构问题: {result['structure_issues']}")
            if result['content_issues'] > 0:
                print(f"    内容问题: {result['content_issues']}")
            if result['file_issues'] > 0:
                print(f"    文件问题: {result['file_issues']}")
            if result['completeness_issues'] > 0:
                print(f"    完整性问题: {result['completeness_issues']}")

    # 保存报告
    report_file = Path('文档') / '技能维护报告.md'
    report_file.parent.mkdir(parents=True, exist_ok=True)

    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(f"# 技能包分析报告\n\n")
        f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"## 概述\n\n")
        f.write(f"- 总技能包数量: {len(analysis_results)}\n")
        f.write(f"- 无问题的技能包: {len([r for r in analysis_results if r['total_issues'] == 0])}\n")
        f.write(f"- 有问题的技能包: {len([r for r in analysis_results if r['total_issues'] > 0])}\n\n")
        f.write(f"## 问题统计\n\n")
        f.write(f"- 结构问题总数: {sum(r['structure_issues'] for r in analysis_results)}\n")
        f.write(f"- 内容问题总数: {sum(r['content_issues'] for r in analysis_results)}\n")
        f.write(f"- 文件问题总数: {sum(r['file_issues'] for r in analysis_results)}\n")
        f.write(f"- 完整性问题总数: {sum(r['completeness_issues'] for r in analysis_results)}\n")
        f.write(f"- 总问题数: {sum(r['total_issues'] for r in analysis_results)}\n\n")
        f.write(f"## 问题最多的技能包\n\n")
        for i, result in enumerate(sorted_results[:5], 1):
            if result['total_issues'] > 0:
                f.write(f"1. {result['name']}: {result['total_issues']} 个问题\n")
        f.write(f"\n## 详细报告\n\n")
        for result in sorted_results:
            if result['total_issues'] > 0:
                f.write(f"### {result['name']}\n\n")
                if result['structure_issues'] > 0:
                    f.write(f"- 结构问题: {result['structure_issues']}\n")
                if result['content_issues'] > 0:
                    f.write(f"- 内容问题: {result['content_issues']}\n")
                if result['file_issues'] > 0:
                    f.write(f"- 文件问题: {result['file_issues']}\n")
                if result['completeness_issues'] > 0:
                    f.write(f"- 完整性问题: {result['completeness_issues']}\n")
                f.write(f"\n")

    print(f"\n报告已保存到: {report_file}")

def main():
    """主函数"""
    print("=" * 60)
    print("技能维护工具")
    print("=" * 60)

    results = []

    for skill_name in SKILL_PACKAGES:
        skill_path = Path('skills') / skill_name
        if skill_path.exists():
            result = analyze_skill(skill_name, skill_path)
            results.append(result)
        else:
            print(f"\n技能包不存在: {skill_name}")

    generate_report(results)

    print("\n" + "=" * 60)
    print("分析完成！")
    print("=" * 60)

if __name__ == '__main__':
    main()
