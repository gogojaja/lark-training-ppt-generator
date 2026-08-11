#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""csv_to_flowchart.py — CSV 节点表 → PPT 流程图（一键转换）

读取 CSV 节点表，应用配色方案，生成流程图 PPT。
可独立使用，也可被 gen_flowchart_branch.py 调用。

== CSV 格式 ==
seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind

字段说明：
  seq          序号（整数，主流程 1-N，分支 41+）
  node_type    节点类型：main=主流程 / branch=分支节点
  content      节点文本内容
  shape        形状：rect=矩形 / diamond=菱形 / circle=圆形 / round_rect=圆角矩形
  width_cm     宽度（cm）
  height_cm    高度（cm）
  bg_color     背景色（6位hex，不含#）
  text_color   字体颜色（6位hex，不含#）
  branch_to    分支目标序号（仅 node_type=main 且有分支时填写）
  branch_label 分支标签文本（如"是"/"否"/"缺失"）
  branch_kind  分支类型：normal=正常分支 / error=异常分支

== 用法 ==
  py -3 csv_to_flowchart.py nodes.csv --out flow.pptx
  py -3 csv_to_flowchart.py nodes.csv --preset green --out flow.pptx
  py -3 csv_to_flowchart.py nodes.csv --preset blue --title "业务流程" --out flow.pptx
"""
import argparse
import csv
import json
import os
import sys
import tempfile
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PRESETS_DIR = os.path.join(SCRIPT_DIR, "..", "skills", "flowchart-skill", "presets")


def load_preset(preset_name):
    """加载配色预设。支持预设名或 JSON 文件路径。"""
    if os.path.isfile(preset_name):
        with open(preset_name, encoding="utf-8") as f:
            return json.load(f)
    path = os.path.join(PRESETS_DIR, preset_name + ".json")
    if os.path.isfile(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    raise SystemExit(f"配色预设 '{preset_name}' 不存在（{path}）")


def csv_to_json(csv_path, preset=None, title=None):
    """读取 CSV 节点表，转换为 gen_flowchart_branch.py 兼容的 JSON 格式。"""
    rows = []
    with open(csv_path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 跳过空行
            if not row.get("seq") or not row.get("content"):
                continue
            rows.append(row)

    if not rows:
        raise SystemExit("CSV 为空或格式不正确。")

    # 加载配色预设
    colors = load_preset(preset) if preset else None

    # 分离主流程和分支节点
    main_nodes = {}
    branch_nodes = {}
    for r in rows:
        seq = int(r["seq"])
        node = {
            "seq": seq,
            "node_type": r.get("node_type", "main").strip(),
            "content": r["content"].strip(),
            "shape": r.get("shape", "rect").strip(),
            "width_cm": float(r.get("width_cm", 5.0)),
            "height_cm": float(r.get("height_cm", 0.6)),
            "bg_color": r.get("bg_color", "C6EFCE").strip(),
            "text_color": r.get("text_color", "006100").strip(),
            "branch_to": r.get("branch_to", "").strip(),
            "branch_label": r.get("branch_label", "").strip(),
            "branch_kind": r.get("branch_kind", "").strip(),
        }
        if node["node_type"] == "branch":
            branch_nodes[seq] = node
        else:
            main_nodes[seq] = node

    # 按序号排序主流程
    sorted_main = sorted(main_nodes.values(), key=lambda x: x["seq"])

    # 应用配色预设覆盖 CSV 颜色
    if colors:
        for node in sorted_main:
            shape = node["shape"]
            if shape == "diamond":
                node["bg_color"] = colors["diamond"]["fill"]
                node["text_color"] = colors["diamond"]["text"]
            elif node["node_type"] == "main":
                node["bg_color"] = colors["main"]["fill"]
                node["text_color"] = colors["main"]["text"]
        for node in branch_nodes.values():
            if node["branch_kind"] == "error":
                node["bg_color"] = colors["error"]["fill"]
                node["text_color"] = colors["error"]["text"]
            else:
                node["bg_color"] = colors["branch"]["fill"]
                node["text_color"] = colors["branch"]["text"]

    # 构建 steps 列表（gen_flowchart_branch.py 语义模式格式）
    steps = []
    for node in sorted_main:
        step = {"text": node["content"]}

        # 应用自定义尺寸到 dim
        step["_w"] = node["width_cm"]
        step["_h"] = node["height_cm"]
        step["_bg"] = node["bg_color"]
        step["_tc"] = node["text_color"]

        # 有分支
        if node["branch_to"]:
            try:
                target_seq = int(node["branch_to"])
                target = branch_nodes.get(target_seq)
                if target:
                    step["branch"] = {
                        "text": target["content"],
                        "label": node["branch_label"],
                        "kind": node["branch_kind"],
                    }
                    step["_br_w"] = target["width_cm"]
                    step["_br_h"] = target["height_cm"]
                    step["_br_bg"] = target["bg_color"]
                    step["_br_tc"] = target["text_color"]
            except (ValueError, KeyError):
                pass

        steps.append(step)

    # 获取标题
    if not title:
        title = os.path.splitext(os.path.basename(csv_path))[0]

    flow = {"title": title, "steps": steps}

    # 如果有配色预设，添加 title 样式
    if colors:
        flow["_title_bg"] = colors.get("title_bg", "1F3864")
        flow["_title_text"] = colors.get("title_text", "FFFFFF")

    return flow


def generate_ppt(flow, out_path, no_conn=False):
    """调用 gen_flowchart_branch.py 生成 PPT。"""
    # 写入临时 JSON
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".json", prefix="fc_")
    try:
        with os.fdopen(tmp_fd, "w", encoding="utf-8") as f:
            json.dump(flow, f, ensure_ascii=False, indent=2)

        # 构建命令
        script = os.path.join(SCRIPT_DIR, "gen_flowchart_branch.py")
        cmd = [sys.executable, script, tmp_path, "--out", out_path]
        if no_conn:
            cmd.append("--no-connectors")

        # 传递维度参数（从 steps 提取默认值）
        if flow.get("steps"):
            s = flow["steps"][0]
            if "_w" in s:
                cmd.extend(["--box-w", str(s["_w"])])
            if "_h" in s:
                cmd.extend(["--box-h", str(s["_h"])])
            if "dim" in flow:
                d = flow["dim"]
                if "step_gap" in d:
                    cmd.extend(["--step-gap", str(d["step_gap"] / 360000)])

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(cmd, capture_output=True, text=True,
                                encoding="utf-8", errors="replace", env=env)
        if result.returncode != 0:
            raise SystemExit(f"gen_flowchart_branch.py 失败:\n{result.stderr}")
        if result.stdout:
            print(result.stdout.strip())
    finally:
        os.unlink(tmp_path)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="CSV 节点表 → PPT 流程图（一键转换）")
    ap.add_argument("csv", help="CSV 节点表文件路径")
    ap.add_argument("--out", default="流程图.pptx", help="输出 PPT 路径")
    ap.add_argument("--preset",
                    help="配色预设名称（green/blue/red/yellow）或预设 JSON 路径")
    ap.add_argument("--title", help="流程图标题（默认从文件名推断）")
    ap.add_argument("--no-connectors", dest="no_conn", action="store_true",
                    help="不生成连接线")
    ap.add_argument("--json-only", action="store_true",
                    help="仅输出 JSON，不生成 PPT（调试用）")
    a = ap.parse_args(argv)

    flow = csv_to_json(a.csv, preset=a.preset, title=a.title)

    if a.json_only:
        out = a.out if a.out.endswith(".json") else a.out + ".json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(flow, f, ensure_ascii=False, indent=2)
        print("JSON 已生成: %s" % out)
    else:
        generate_ppt(flow, a.out, no_conn=a.no_conn)


if __name__ == "__main__":
    sys.exit(main())
