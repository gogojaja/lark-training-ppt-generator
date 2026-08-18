#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
table-skill 启动脚本
用于快速生成表格页
"""

import sys
import os
import json

# 添加生成脚本目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', '生成脚本'))

from gen_table import create_table_slide
from pptx import Presentation
from pptx.util import Inches

def main():
    """主函数"""
    print("=" * 60)
    print("table-skill 表格页生成器")
    print("=" * 60)

    # 检查是否有 JSON 文件
    if len(sys.argv) < 2:
        print("\n使用方法：")
        print("  python table_skill.py <json_file> [template] [preset] [output_file]")
        print("\n参数说明：")
        print("  json_file  : JSON 数据文件路径（必需）")
        print("  template   : 模板类型（可选，默认：standard）")
        print("  preset     : 配色方案（可选，默认：blue）")
        print("  output_file: 输出文件名（可选，默认：表格.pptx）")
        print("\n模板类型：")
        print("  standard  - 标准表格")
        print("  zebra     - 斑马纹表格")
        print("  icon      - 带图标表格")
        print("\n配色方案：")
        print("  blue   - 专业蓝")
        print("  gray   - 商务灰")
        sys.exit(1)

    # 读取参数
    json_file = sys.argv[1]
    template = sys.argv[2] if len(sys.argv) > 2 else 'standard'
    preset = sys.argv[3] if len(sys.argv) > 3 else 'blue'
    output_file = sys.argv[4] if len(sys.argv) > 4 else '表格.pptx'

    # 检查 JSON 文件是否存在
    if not os.path.exists(json_file):
        print(f"\n错误：找不到文件 {json_file}")
        sys.exit(1)

    # 读取 JSON 数据
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"\n错误：读取 JSON 文件失败 - {e}")
        sys.exit(1)

    # 检查 JSON 数据格式
    required_fields = ['title', 'headers', 'rows']
    for field in required_fields:
        if field not in data:
            print(f"\n错误：JSON 缺少必需字段 - {field}")
            sys.exit(1)

    # 创建 PPT
    print(f"\n正在生成表格页...")
    print(f"  标题: {data['title']}")
    print(f"  模板: {template}")
    print(f"  配色: {preset}")
    print(f"  数据行数: {len(data['rows'])}")
    print(f"  数据列数: {len(data['headers'])}")

    try:
        pres = Presentation()
        pres.slide_width = Inches(9.5)
        pres.slide_height = Inches(6.5)

        create_table_slide(pres, data, template, preset)

        # 保存
        output_path = output_file
        pres.save(output_path)

        print(f"\n成功！表格页已保存到：{output_path}")
        print(f"文件大小: {os.path.getsize(output_path) / 1024:.2f} KB")

    except Exception as e:
        print(f"\n错误：生成表格页失败 - {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("table-skill 表格页生成完成")
    print("=" * 60)

if __name__ == '__main__':
    main()
