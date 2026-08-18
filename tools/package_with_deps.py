#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自包含打包脚本（package_with_deps.py）

方案 A：源码保持单源，交付时把 5 个外部 lark-* 技能内嵌为 `_lark/` 副本，
并改写 SKILL.md / 交接文档.md 中 ../lark-* 引用为项目内路径，产出自包含目录。

范围治理边界（2026-08-18 固，见 范围初定义说明书.md §7）：
- 本项目工作范围限定于本仓库，lark-* 5 技能由独立仓库管理维护；
- 打包时仅将同级目录已克隆的 lark-* 独立技能副本内嵌进产物，不修改其源码；
- 禁止硬编码平台绝对路径（如 D:\trae\...），外部技能根目录可经 --lark-root 指定，
  缺省为本项目同级目录。

用法（在项目根运行）:
    py -3 tools/package_with_deps.py                      # 打包到 dist/
    py -3 tools/package_with_deps.py --out D:/out/pkg     # 指定输出目录
    py -3 tools/package_with_deps.py --lark-root D:/tools  # 指定外部技能根目录

输出:
    dist/制度手册转宣讲PPT工作流_v4.1.0_自包含/       （目录，含内嵌 _lark/lark-*）
    dist/制度手册转宣讲PPT工作流_v4.1.0_自包含.zip

注意:
    - 本脚本不改动源码目录，只生成打包产物；
    - 源码 SKILL.md 仍引用 ../lark-*（单源），内嵌改写仅发生在打包副本内。
"""
import argparse
import os
import re
import shutil
import zipfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT = os.path.dirname(PROJECT_ROOT)  # 外部技能同级目录（可经 --lark-root 覆盖）
VERSION = "v4.1.0"
PACK_NAME = "制度手册转宣讲PPT工作流"
EMBED_DIR = "_lark"  # 打包产物内 lark-* 副本目录名

# 5 个外部 lark 技能（独立仓库管理，需先 clone 到 --lark-root 目录）
LARK_DEPS = ["lark-shared", "lark-slides", "lark-doc", "lark-wiki", "lark-base"]

# 参与打包的项目内容（相对项目根）
INCLUDE_ITEMS = ["SKILL.md", "交接文档.md", "生成脚本", "素材", "台账"]


LARK_ROOT = PARENT


def lark_source_dir(name: str) -> str:
    return os.path.join(LARK_ROOT, name)


def rewrite_embed_dir(m: re.Match) -> str:
    """把 `../lark-xxx` 替换为 `_lark/lark-xxx`（保留链接文本不变）"""
    return f"{EMBED_DIR}/lark-{m.group(1)}"


def rewrite_text(text: str) -> str:
    # 1. 链接/引用 ../lark-xxx
    text = re.sub(r"\.\./lark-(shared|slides|doc|wiki|base)", rewrite_embed_dir, text)
    # 2. 绝对路径 D:/trae/lark-xxx 或 D:\trae\lark-xxx
    for name in ("shared", "slides", "doc", "wiki", "base"):
        for sep in ("/", "\\"):
            text = text.replace(f"D:{sep}trae{sep}lark-{name}", f"{EMBED_DIR}/lark-{name}")
    return text


def build_package(out_dir: str) -> str:
    out_dir = os.path.abspath(out_dir)
    root = os.path.join(out_dir, f"{PACK_NAME}_{VERSION}_自包含")
    if os.path.exists(root):
        shutil.rmtree(root)
    os.makedirs(root)

    # 1. 拷贝项目内容
    for item in INCLUDE_ITEMS:
        src = os.path.join(PROJECT_ROOT, item)
        dst = os.path.join(root, item)
        if not os.path.exists(src):
            print(f"  [跳过] {item} 不存在")
            continue
        if os.path.isdir(src):
            shutil.copytree(src, dst)
        else:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)

    # 2. 内嵌 lark-* 依赖副本
    embed_root = os.path.join(root, EMBED_DIR)
    os.makedirs(embed_root, exist_ok=True)
    for name in LARK_DEPS:
        src = lark_source_dir(name)
        if not os.path.isdir(src):
            print(f"  [警告] 未找到外部技能 {name}（预期 {src}），跳过内嵌")
            continue
        shutil.copytree(src, os.path.join(embed_root, name), dirs_exist_ok=True)
        print(f"  [内嵌] {name} -> {EMBED_DIR}/{name}")

    # 3. 改写打包副本内的引用路径
    for rel in ("SKILL.md", "交接文档.md"):
        fp = os.path.join(root, rel)
        if not os.path.exists(fp):
            continue
        with open(fp, encoding="utf-8") as fh:
            text = rewrite_text(fh.read())
        with open(fp, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"  [改写] {rel} 引用 -> {EMBED_DIR}/")

    # 4. 打包 zip
    zip_path = root + ".zip"
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder, _sub, files in os.walk(root):
            for f in files:
                fp = os.path.join(folder, f)
                arc = os.path.relpath(fp, os.path.dirname(root))
                zf.write(fp, arc)
    print(f"  [打包完成] {zip_path}")
    return zip_path


def main():
    ap = argparse.ArgumentParser(description="自包含打包脚本")
    ap.add_argument("--out", default="dist", help="打包输出目录（相对项目根或绝对路径）")
    ap.add_argument("--lark-root", default=PARENT, help="外部 lark-* 技能根目录（缺省=项目同级目录）")
    args = ap.parse_args()
    global LARK_ROOT
    LARK_ROOT = os.path.abspath(args.lark_root)
    out_dir = args.out if os.path.isabs(args.out) else os.path.join(PROJECT_ROOT, args.out)
    zip_path = build_package(out_dir)
    print("\n打包完成，产物：", zip_path)
    print("（源码 SKILL.md 仍指向 ../lark-*，保持单源未改动）")


if __name__ == "__main__":
    main()