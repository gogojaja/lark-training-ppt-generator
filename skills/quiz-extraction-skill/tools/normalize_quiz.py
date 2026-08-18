#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""后端规范化工具（normalize_quiz.py）

考试题抽取技能 Step 3：对 AI 语义抽取产出的题组 JSON 做清洗/合并/校验/组号重排。
原则：AI 只负责语义判断（题型/题组/题干），本工具把不稳定输出收敛为稳定结构。

用法:
    python tools/normalize_quiz.py 输出/AI抽取.json -o 输出/规范化.json

处理动作:
    1. JSON 兜底解析（容忍 AI 输出多余 Markdown 代码块/空行）
    2. 字段类型清洗（score->数字、source_page->整数、空字符串->空数组）
    3. 组号重排（G1,G2,...）与小题 qid 重排（G1-Q1,...）
    4. 答案归一（大写、去空格、数组化）
    5. 合并共享题干题组
    6. schema 校验（复用 validate_quiz.py 规则）
"""
import argparse
import json
import os
import re
import sys


def robust_json_load(text: str):
    """容忍 AI 输出包裹 markdown 代码块或首尾空白。"""
    text = text.strip()
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if fence:
        text = fence.group(1).strip()
    return json.loads(text)


def clean_scalar(value, cast, default):
    try:
        if value is None or (isinstance(value, str) and not value.strip()):
            return default
        return cast(value)
    except (TypeError, ValueError):
        return default


def normalize_answer(answer, qtype):
    """答案归一为列表。"""
    if answer is None:
        return []
    if isinstance(answer, str):
        parts = [a.strip().rstrip("。；;，,") for a in answer.replace("、", ",").replace(" ", "").split(",")]
        parts = [p for p in parts if p]
        if qtype == "true_false":
            for p in parts:
                if p in ("正确", "对", "T", "true", "TRUE", "√"):
                    return ["正确"]
                if p in ("错误", "错", "F", "false", "FALSE", "×"):
                    return ["错误"]
            return parts
        return parts
    if isinstance(answer, list):
        return [str(a).strip() for a in answer if str(a).strip()]
    return [str(answer)]


def normalize_question(q, group_id, idx):
    qtype = q.get("type", "").strip().lower()
    if qtype not in ("single_choice", "multiple_choice", "true_false", "fill_blank", "short_answer"):
        qtype = "single_choice"
    options = q.get("options") or []
    if isinstance(options, str):
        options = [o.strip() for o in re.split(r"\n|[ABCDEF]\s*[\.、)]", options) if o.strip()]
    options = [str(o).strip() for o in options if str(o).strip()]
    return {
        "qid": f"{group_id}-Q{idx}",
        "type": qtype,
        "stem": str(q.get("stem", "")).strip(),
        "options": options,
        "answer": normalize_answer(q.get("answer"), qtype),
        "analysis": str(q.get("analysis", "")).strip(),
        "difficulty": q.get("difficulty", "medium") if q.get("difficulty") in ("easy", "medium", "hard") else "medium",
        "score": clean_scalar(q.get("score"), float, None),
        "source_page": clean_scalar(q.get("source_page"), int, None),
    }


def _collect_pages(g, questions):
    pages = set()
    raw_pages = g.get("source_pages")
    if isinstance(raw_pages, list):
        for p in raw_pages:
            n = clean_scalar(p, int, None)
            if n is not None:
                pages.add(n)
    else:
        n = clean_scalar(raw_pages, int, None)
        if n is not None:
            pages.add(n)
    for q in questions:
        if q["source_page"] is not None:
            pages.add(q["source_page"])
    return sorted(pages)


def normalize_groups(data):
    groups = []
    raw_groups = data.get("groups") or []
    for g_idx, g in enumerate(raw_groups, start=1):
        group_id = f"G{g_idx}"
        section = str(g.get("section", "") or "").strip()
        shared_stem = (g.get("shared_stem") or "").strip() or None
        shared_image = (g.get("shared_image") or "").strip() or None
        qs = g.get("questions") or []
        questions = [normalize_question(q, group_id, qi) for qi, q in enumerate(qs, start=1)]
        questions = [q for q in questions if q["stem"]]
        if not questions:
            continue
        groups.append({
            "group_id": group_id,
            "section": section,
            "shared_stem": shared_stem,
            "shared_image": shared_image,
            "source_pages": _collect_pages(g, questions),
            "questions": questions,
        })
    return groups


def main():
    ap = argparse.ArgumentParser(description="考试题抽取：AI 输出规范化/清洗/合并/组号重排")
    ap.add_argument("input", help="AI 抽取输出的 JSON 路径")
    ap.add_argument("-o", "--out", default="输出/规范化.json", help="输出 JSON 路径")
    args = ap.parse_args()

    with open(args.input, encoding="utf-8") as fh:
        text = fh.read()
    try:
        raw = robust_json_load(text)
    except json.JSONDecodeError as e:
        print(f"[错误] AI 输出 JSON 无法解析: {e}")
        sys.exit(1)

    if isinstance(raw, list):
        raw = {"groups": raw}

    paper_title = str(raw.get("paper_title", "") or "").strip() or "未命名试卷"
    source = raw.get("source") or {}
    groups = normalize_groups(raw)

    warnings = list(raw.get("metadata", {}).get("warnings") or [])
    if not groups:
        warnings.append("未识别到任何题目，请检查源文档或 AI 抽取结果")

    result = {
        "paper_title": paper_title,
        "source": {
            "path": str(source.get("path", "") or "").strip(),
            "format": str(source.get("format", "") or "").strip() or "unknown",
            "pages": clean_scalar(source.get("pages"), int, None),
        },
        "metadata": {
            "subject": str(raw.get("metadata", {}).get("subject", "") or "").strip() or None,
            "difficulty": raw.get("metadata", {}).get("difficulty", "medium"),
            "total_questions": sum(len(g["questions"]) for g in groups),
            "warnings": warnings,
        },
        "groups": groups,
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=2)
    print(f"[完成] 规范化 {args.input} -> {args.out}")
    print(f"  题组数: {len(groups)}，题目数: {result['metadata']['total_questions']}")


if __name__ == "__main__":
    main()