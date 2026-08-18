#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Schema 校验工具（validate_quiz.py）

考试题抽取技能验收门禁：对规范化后的题组 JSON 做结构校验。
校验不通过返回非 0 退出码并打印问题清单。

用法:
    python tools/validate_quiz.py 输出/规范化.json

通过条件（全部满足）:
    - JSON 可解析，顶层含 paper_title / source / groups
    - 每个 group 含 group_id / section / questions(>=1)
    - 每道题 qid / type(枚举) / stem 非空
    - type 枚举: single_choice / multiple_choice / true_false / fill_blank / short_answer
    - 选择题: options 非空，且 answer 均落在选项内
    - 判断题: answer 为 正确/错误
    - qid 全局唯一
"""
import argparse
import json
import re
import sys

VALID_TYPES = {"single_choice", "multiple_choice", "true_false", "fill_blank", "short_answer"}


def option_labels(options):
    labels = set()
    for opt in options:
        m = re.match(r"^\s*([A-Fa-f])", opt)
        if m:
            labels.add(m.group(1).upper())
    return labels


def validate(data):
    errors, warnings = [], []

    for field in ("paper_title", "source", "groups"):
        if field not in data:
            errors.append(f"顶层缺少字段: {field}")
    if not data.get("groups"):
        errors.append("groups 为空")

    seen_qids = set()
    for g in data.get("groups", []):
        if "group_id" not in g:
            errors.append("存在缺少 group_id 的题组")
        if not g.get("questions"):
            errors.append(f"{g.get('group_id','?')} 缺少题目")
            continue
        for q in g["questions"]:
            qid = q.get("qid")
            if not qid:
                errors.append(f"{g.get('group_id','?')} 存在无 qid 的题目")
            elif qid in seen_qids:
                errors.append(f"qid 重复: {qid}")
            else:
                seen_qids.add(qid)
            qtype = q.get("type")
            if qtype not in VALID_TYPES:
                errors.append(f"{qid} 非法题型: {qtype}")
                continue
            if not str(q.get("stem", "")).strip():
                errors.append(f"{qid} 题干为空")
            if qtype in ("single_choice", "multiple_choice"):
                options = q.get("options") or []
                if not options:
                    errors.append(f"{qid} 选择题缺少选项")
                labels = option_labels(options)
                if qtype == "single_choice" and len(q.get("answer") or []) != 1:
                    errors.append(f"{qid} 单选题答案数量应为 1")
                if qtype == "multiple_choice" and len(q.get("answer") or []) < 1:
                    errors.append(f"{qid} 多选题答案不能为空")
                for ans in q.get("answer") or []:
                    if labels and ans.strip().upper() not in labels:
                        errors.append(f"{qid} 答案 {ans} 不在选项内")
            if qtype == "true_false":
                if not q.get("answer"):
                    errors.append(f"{qid} 判断题缺少答案")
                elif any(a not in ("正确", "错误") for a in q["answer"]):
                    errors.append(f"{qid} 判断题答案应为 正确/错误")
            if qtype == "fill_blank":
                if not q.get("answer"):
                    warnings.append(f"{qid} 填空题无答案（可人工补充）")
            if qtype == "short_answer":
                if not q.get("answer"):
                    warnings.append(f"{qid} 简答题无参考答案要点")

    return errors, warnings


def main():
    ap = argparse.ArgumentParser(description="考试题抽取：schema 校验")
    ap.add_argument("input", help="规范化 JSON 路径")
    ap.add_argument("--strict", action="store_true", help="warning 也视为失败")
    args = ap.parse_args()

    with open(args.input, encoding="utf-8") as fh:
        data = json.load(fh)

    errors, warnings = validate(data)
    for e in errors:
        print(f"[错误] {e}")
    for w in warnings:
        print(f"[警告] {w}")

    if errors or (args.strict and warnings):
        print(f"[失败] 校验未通过（{len(errors)} 错误，{len(warnings)} 警告）")
        sys.exit(1)
    print(f"[通过] 校验通过（{len(data.get('groups', []))} 题组，"
          f"{sum(len(g['questions']) for g in data.get('groups', []))} 题）")
    sys.exit(0)


if __name__ == "__main__":
    main()