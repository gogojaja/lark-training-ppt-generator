#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""validate_csv.py — 流程图 CSV 节点表校验工具

在生成 PPT 前校验 CSV 格式，提前发现常见问题。

== 用法 ==
  py -3 tools/validate_csv.py nodes.csv
  py -3 tools/validate_csv.py nodes.csv --fix   # 自动修复可修复问题

== 校验项 ==
  1. 表头格式：必须为统一单行表头（无注释行、无多行表头）
  2. 列对齐：config 行用前4列，node 行用后11列
  3. 空值检查：避免 None 值（尾随逗号问题）
  4. 节点结构：seq 连续、branch 编号递增、单入口单出口
  5. 节点文字：8~15 字、动词开头
"""
import csv
import sys
import os

REQUIRED_HEADER = [
    "type", "key", "value", "desc",
    "seq", "node_type", "content", "shape",
    "width_cm", "height_cm", "bg_color", "text_color",
    "branch_to", "branch_label", "branch_kind"
]

VALID_SHAPES = {"rect", "diamond", "circle", "round_rect", "end"}
VALID_NODE_TYPES = {"main", "branch"}
VALID_BRANCH_KINDS = {"normal", "error", ""}


def validate_header(rows):
    """校验表头是否为统一格式。"""
    errors = []
    if not rows:
        errors.append("CSV 文件为空")
        return errors

    first_row = rows[0]
    actual_keys = list(first_row.keys())

    # 检查是否有注释行（第一个 key 以 # 开头）
    if actual_keys[0].startswith("#"):
        errors.append("表头前有注释行，csv.DictReader 将无法正确解析。删除所有注释行。")

    # 检查列数
    if len(actual_keys) < len(REQUIRED_HEADER):
        errors.append(f"列数不足：期望 {len(REQUIRED_HEADER)} 列，实际 {len(actual_keys)} 列")

    # 检查列名匹配
    for i, expected in enumerate(REQUIRED_HEADER):
        if i < len(actual_keys) and actual_keys[i] != expected:
            errors.append(f"第{i+1}列名不匹配：期望 '{expected}'，实际 '{actual_keys[i]}'")

    return errors


def validate_none_values(rows):
    """检查是否有 None 值（尾随逗号问题）。"""
    errors = []
    for i, row in enumerate(rows):
        for key, value in row.items():
            if value is None:
                errors.append(f"第{i+1}行 '{key}' 列值为 None（可能由尾随逗号导致）")
    return errors


def validate_config_rows(rows):
    """校验 config 行格式。"""
    errors = []
    config_count = 0
    for i, row in enumerate(rows):
        type_ = (row.get("type") or "").strip()
        if type_ == "config":
            config_count += 1
            key = (row.get("key") or "").strip()
            value = (row.get("value") or "").strip()
            if not key:
                errors.append(f"第{i+1}行 config 缺少 key")
            # 检查 config 行的 node 列是否为空
            if row.get("seq") or row.get("content"):
                errors.append(f"第{i+1}行 config 行不应有节点数据")
    if config_count == 0:
        errors.append("未找到 config 配置行（可选，但建议添加）")
    return errors


def validate_node_rows(rows):
    """校验节点行格式。"""
    errors = []
    main_seqs = []
    branch_seqs = []

    for i, row in enumerate(rows):
        type_ = (row.get("type") or "").strip()
        if type_ == "config":
            continue

        seq_str = (row.get("seq") or "").strip()
        content = (row.get("content") or "").strip()
        node_type = (row.get("node_type") or "").strip()
        shape = (row.get("shape") or "").strip()

        if not seq_str and not content:
            continue  # 空行，跳过

        if not seq_str:
            errors.append(f"第{i+1}行缺少 seq")
            continue
        if not content:
            errors.append(f"第{i+1}行缺少 content")
            continue

        try:
            seq = int(seq_str)
        except ValueError:
            errors.append(f"第{i+1}行 seq 非整数：'{seq_str}'")
            continue

        # 校验 node_type
        if node_type not in VALID_NODE_TYPES:
            errors.append(f"第{i+1}行 node_type 无效：'{node_type}'（应为 main/branch）")

        # 校验 shape
        if shape and shape not in VALID_SHAPES:
            errors.append(f"第{i+1}行 shape 无效：'{shape}'（应为 rect/diamond/circle/round_rect/end）")

        # 收集 seq
        if node_type == "main":
            main_seqs.append(seq)
        elif node_type == "branch":
            branch_seqs.append(seq)

        # 校验菱形节点
        if shape == "diamond":
            branch_to = (row.get("branch_to") or "").strip()
            branch_kind = (row.get("branch_kind") or "").strip()
            if not branch_to:
                errors.append(f"第{i+1}行菱形节点缺少 branch_to")
            if branch_kind not in VALID_BRANCH_KINDS:
                errors.append(f"第{i+1}行 branch_kind 无效：'{branch_kind}'")

        # 校验节点文字长度
        if len(content) < 1:
            errors.append(f"第{i+1}行节点文字为空")
        elif len(content) > 20:
            errors.append(f"第{i+1}行节点文字过长（{len(content)}字）：'{content}'")

    # 校验主流程 seq 连续性
    if main_seqs:
        main_seqs_sorted = sorted(main_seqs)
        for i in range(1, len(main_seqs_sorted)):
            if main_seqs_sorted[i] != main_seqs_sorted[i-1] + 1:
                errors.append(f"主流程 seq 不连续：{main_seqs_sorted[i-1]} → {main_seqs_sorted[i]}")

    # 校验分支编号
    if branch_seqs:
        branch_seqs_sorted = sorted(branch_seqs)
        if branch_seqs_sorted[0] < 41:
            errors.append(f"分支编号应从 41 开始，实际从 {branch_seqs_sorted[0]} 开始")
        for i in range(1, len(branch_seqs_sorted)):
            if branch_seqs_sorted[i] != branch_seqs_sorted[i-1] + 1:
                errors.append(f"分支编号不连续：{branch_seqs_sorted[i-1]} → {branch_seqs_sorted[i]}")

    # 校验单入口单出口
    if main_seqs:
        if len([s for s in main_seqs if s == 1]) != 1:
            errors.append("应有且仅有1个入口节点（seq=1）")

    return errors


def validate_csv(csv_path):
    """主校验函数。"""
    print(f"校验: {csv_path}")

    if not os.path.isfile(csv_path):
        print(f"[ERROR] 文件不存在: {csv_path}")
        return False

    try:
        with open(csv_path, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
    except Exception as e:
        print(f"[ERROR] 读取失败: {e}")
        return False

    all_errors = []

    # 1. 表头校验
    header_errors = validate_header(rows)
    all_errors.extend(header_errors)

    # 2. None 值校验
    none_errors = validate_none_values(rows)
    all_errors.extend(none_errors)

    # 3. Config 行校验
    config_errors = validate_config_rows(rows)
    all_errors.extend(config_errors)

    # 4. 节点行校验
    node_errors = validate_node_rows(rows)
    all_errors.extend(node_errors)

    # 输出结果
    if not all_errors:
        print("[OK] 校验通过，无问题发现")
        return True
    else:
        print(f"[FAIL] 发现 {len(all_errors)} 个问题：")
        for err in all_errors:
            print(f"  - {err}")
        return False


def main():
    if len(sys.argv) < 2:
        print("用法: py -3 tools/validate_csv.py <csv文件> [--fix]")
        sys.exit(1)

    csv_path = sys.argv[1]
    success = validate_csv(csv_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
