#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_to_flow_csv.py - 从 Word 文档生成流程图 CSV 节点表（纯规则启发式草稿）

不依赖任何第三方库与外部模型，仅用 Python 标准库 + docx 的 XML 解析。
从 Word 文档段落中按关键字启发式识别"处理步骤/判断节点"，
生成与 flowchart-skill 相同 schema 的 CSV 节点表草稿。

用法：
    py -3 tools/docx_to_flow_csv.py <docx文件> --out nodes.csv
    py -3 tools/docx_to_flow_csv.py <docx文件> --out nodes.csv --preset green
    py -3 tools/docx_to_flow_csv.py <docx文件> --json-only --out flow.json   # 调试

说明：
    输出为【草稿】，仅供后续"手调整"起步，需人工核对流程完整性、
    判断分支归属、文字精简后再交由 csv_to_flowchart.py 生成 PPT。

== CSV 输出字段 ==
seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind
"""

import argparse
import csv
import json
import os
import re
import sys
import zipfile


# ---------------- docx 文本提取 ----------------

def extract_text(elem_bytes):
    """从段落字节中提取文本（跨 <w:t>，合并）"""
    texts = re.findall(rb'<w:t[^>]*>([^<]*)</w:t>', elem_bytes)
    return b''.join(texts).decode('utf-8', errors='replace').strip()


def extract_paras_from_docx(docx_path):
    """按出现顺序返回 [(text, is_heading, outline_level), ...]"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        doc_xml = z.read('word/document.xml')

    body_start = doc_xml.find(b'<w:body>')
    body_end = doc_xml.find(b'</w:body>')
    if body_start < 0 or body_end < 0:
        body = doc_xml
    else:
        body = doc_xml[body_start:body_end + 8]

    # 段落边界
    para_starts = []
    pos = 0
    while True:
        idx = body.find(b'<w:p ', pos)
        if idx < 0:
            idx = body.find(b'<w:p>', pos)
        if idx < 0:
            break
        para_starts.append(idx)
        pos = idx + 4

    if not para_starts:
        return []

    paras = []
    for i, start in enumerate(para_starts):
        end = para_starts[i + 1] if i + 1 < len(para_starts) else len(body)
        elem = body[start:end]
        text = extract_text(elem)
        if not text:
            continue
        # 标题检测
        is_heading = False
        outlvl = None
        m_lvl = re.search(rb'outlineLvl w:val="(\d+)"', elem)
        if m_lvl:
            is_heading = True
            outlvl = int(m_lvl.group(1).decode())
        elif re.search(rb'<w:pStyle w:val="Heading\d+"', elem):
            is_heading = True
            m_hm = re.search(rb'Heading(\d)', elem)
            outlvl = int(m_hm.group(1)) - 1 if m_hm else 0
        paras.append((text, is_heading, outlvl))
    return paras


# ---------------- 规则识别 ----------------

# 动作动词（句子开头命中即为处理步骤）
ACTION_VERBS = [
    "客户", "用户", "柜员", "操作", "经办", "审核", "审批", "办理", "输入", "填写",
    "选择", "点击", "提交", "确认", "打印", "查询", "核对", "校验", "核实", "检查",
    "登记", "录入", "扫描", "读取", "识别", "认证", "验证", "拍摄", "上传", "下载",
    "发送", "接收", "生成", "保存", "上传", "授权", "开户", "签约", "挂失", "冻结",
    "付款", "转账", "结算", "审核", "复核", "授权", "审批", "通知", "联系", "联系",
    "系统", "平台", "登录", "退出", "到达", "前往", "等待", "开始", "完成", "结束",
]

# 判断关键词（命中即判断节点 diamond）
JUDGE_KEYWORDS = [
    "是否", "若", "如", "如果", "否则", "需要", "须", "是否通过", "校验通过",
    "校验失败", "审核通过", "审核不通过", "存在", "不存在", "符合", "不符合", "异常",
    "失败", "成功", "缺", "无", "有效", "无效", "通过", "不通过", "确认无误",
    "大于", "小于", "等于", "超过", "不足", "正常", "没问题", "有误",
]

# 分支标签常见词
BRANCH_ERR_HINTS = ["不通过", "失败", "异常", "缺失", "无效", "不符合", "不足", "超时"]
BRANCH_OK_HINTS = ["通过", "成功", "有效", "符合", "正常", "无误"]

# 冗余词（精简时剔除）
REDUNDANT = [
    "首先", "然后", "之后", "接着", "最后", "同时", "并", "并且", "以及", "如",
    "例如", "等", "相关", "相应", "进行", "完成", "及", "与", "将", "需要",
]

# 明确的判断句式："若/如果/当...则/否则"
COND_PATTERN = re.compile(r'(若|如果|当|如).{0,12}?(则|就|应|需|可|直接|否则|需)')
ELSE_PATTERN = re.compile(r'否则|否则如果|反之|不然')


def is_possible_title(text):
    """判断是否像标题（短、无句号、像章节名）"""
    if len(text) > 30:
        return False
    if any(p in text for p in "。，；："):
        return False
    return True


def split_sentences(text):
    """按中文/英文标点切分为短句，去空"""
    parts = re.split(r'[。；;\n]', text)
    out = []
    for p in parts:
        p = p.strip()
        if p:
            out.append(p)
    return out


def starts_with_verb(s):
    return any(s.startswith(v) for v in ACTION_VERBS) or any(
        s.startswith(v) for v in ["点击", "选择", "输入", "提交", "确认", "打印", "填写", "扫描"])


def contains_judge(s):
    return any(kw in s for kw in JUDGE_KEYWORDS)


def compress(s, max_len=15):
    """精简节点文字：去冗余词、截断"""
    text = s.strip()
    for r in REDUNDANT:
        text = text.replace(r, '')
    # 去掉引号括号内过长内容
    text = re.sub(r'[（(][^）)]{6,}[）)]', '', text)
    text = re.sub(r'[“”"\']', '', text).strip()
    text = re.sub(r'\s+', '', text)
    if len(text) > max_len:
        text = text[:max_len]
    return text or '流程步骤'


# ---------------- 主流程 ----------------

def build_csv(docx_path, preset="green"):
    paras = extract_paras_from_docx(docx_path)

    nodes = []          # 识别出的流程节点 (dict)
    branch_nodes = []   # 判断分支节点

    seq_main = 0
    seq_branch = 40

    # 判断节点栈（最近的有分支的判断，等待下一节点作为正常分支目标）
    pending_judge = None

    for text, is_heading, outlvl in paras:
        if is_heading:
            # 标题跳过（不作为节点），但章节标题可作为后续上下文（暂略）
            continue
        if is_possible_title(text) and len(text) <= 12 and not contains_judge(text) \
                and not starts_with_verb(text):
            # 短标题类文本（如"一、业务规则"）跳过
            continue

        sents = split_sentences(text)
        for s in sents:
            s = s.strip()
            if not s:
                continue
            is_judge = contains_judge(s)
            is_act = starts_with_verb(s) or (len(s) <= 20)

            if not (is_judge or is_act):
                continue

            content = compress(s)
            if is_judge:
                seq_main += 1
                node = {
                    "seq": seq_main,
                    "node_type": "main",
                    "content": content,
                    "shape": "diamond",
                    "cols": (4.5, 1.0),
                    "branch_to": "",
                    "branch_label": "",
                    "branch_kind": "",
                }
                nodes.append(node)
                # 尝试为判断生成分支
                if ELSE_PATTERN.search(s):
                    seq_branch += 1
                    branch_nodes.append({
                        "ref_seq": seq_main,
                        "seq": seq_branch,
                        "content": "其他处理",
                        "kind": "error",
                    })
                pending_judge = seq_main
            else:
                seq_main += 1
                node = {
                    "seq": seq_main,
                    "node_type": "main",
                    "content": content,
                    "shape": "rect",
                    "cols": (5.0, 0.6),
                    "branch_to": "",
                    "branch_label": "",
                    "branch_kind": "",
                }
                nodes.append(node)
                pending_judge = None

    return nodes, branch_nodes


def write_csv(nodes, branch_nodes, out_path, preset="green", title_hint=""):
    """把节点与分支写成标准 CSV"""
    need = preset  # 触发 import 而不使用会告警，这里仅透传给读取方（无实际影响）
    header = "seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind"
    rows = [header]

    main_seq = sorted((n["seq"] for n in nodes))
    # 为判断节点补分支（跳转）
    bnode_by_ref = {}
    for b in branch_nodes:
        bnode_by_ref.setdefault(b["ref_seq"], []).append(b)

    for n in nodes:
        if n["shape"] == "diamond" and n["seq"] in bnode_by_ref:
            for b in bnode_by_ref[n["seq"]]:
                n["branch_to"] = str(b["seq"])
                n["branch_label"] = "不通过" if "err" in b["kind"].lower() else "正常"
                n["branch_kind"] = "error" if "err" in b["kind"].lower() else "normal"
        w, h = n["cols"]
        rows.append(
            f"{n['seq']},{n['node_type']},{n['content']},"
            f"{n['shape']},{w},{h},C6EFCE,006100,"
            f"{n['branch_to']},{n['branch_label']},{n['branch_kind']}"
        )

    for b in branch_nodes:
        w, h = 5.0, 0.6
        bg = "FCE4EC" if "err" in b["kind"].lower() else "DDEBF7"
        tc = "C00000" if "err" in b["kind"].lower() else "1F3864"
        rows.append(
            f"{b['seq']},branch,{b['content']},"
            f"rect,{w},{h},{bg},{tc},,,"
        )

    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(rows) + "\n")

    return len(nodes), len(branch_nodes)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("docx", help="输入 Word 文档 (.docx)")
    ap.add_argument("--out", default="flowchart_nodes.csv", help="输出 CSV 路径")
    ap.add_argument("--preset", default="green", help="配色预设（仅标注，不影响生成）")
    ap.add_argument("--json-only", action="store_true", help="仅输出 JSON 草稿（调试）")
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print(f"文件不存在: {args.docx}")
        sys.exit(1)

    try:
        nodes, branches = build_csv(args.docx)
    except zipfile.BadZipFile:
        print("无法读取 docx（不是有效的 Word 文件）")
        sys.exit(1)

    if args.json_only:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump({"nodes": nodes, "branches": branches}, f, ensure_ascii=False, indent=2)
        print(f"JSON 草稿已写入: {args.out}")
        sys.exit(0)

    if not nodes:
        print("未识别到流程节点，请检查文档内容（可能需要先拆分为业务章节）。")
        sys.exit(1)

    n_nodes, n_branch = write_csv(nodes, branches, args.out, args.preset)
    print(f"已生成 CSV 草稿: {args.out}")
    print(f"  主流程节点: {n_nodes}  分支节点: {n_branch}")
    print("注意：此为规则启发式草稿，请人工核对流程完整性、判断分支与文字精简后使用。")


if __name__ == "__main__":
    main()
