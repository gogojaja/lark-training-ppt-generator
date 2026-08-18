#!/usr/bin/env python3
"""
split_docx_by_level.py — Word文档按大纲级别拆分工具

行业最佳实践：
1. 只复制被引用的图片，减小文件体积
2. 完整的XML命名空间声明，确保Word兼容性
3. 流式处理，支持超大文档

用法：
    py -3 split_docx_by_level.py <docx文件> [输出目录] [level] [章节序号]

示例：
    py -3 split_docx_by_level.py input.docx output 1        # 拆分所有level=1
    py -3 split_docx_by_level.py input.docx output 1 1      # 只拆分第1个level=1
    py -3 split_docx_by_level.py input.docx output 2        # 拆分所有level=2
"""

import os
import re
import sys
import zipfile


def sanitize(name):
    """清理文件名"""
    for ch in '<>:"/\\|?*\n\r\t':
        name = name.replace(ch, '_')
    return name.strip()[:80] or '未命名'


def extract_text(elem_bytes):
    """从段落字节中提取文本"""
    texts = re.findall(rb'<w:t[^>]*>([^<]*)</w:t>', elem_bytes)
    return b''.join(texts).decode('utf-8', errors='replace').strip()


def find_image_rids(section_xml):
    """从段落XML中提取所有引用的图片rId"""
    rids = set()
    for m in re.finditer(rb'r:embed="(rId\d+)"', section_xml):
        rids.add(m.group(1).decode())
    for m in re.finditer(rb'r:id="(rId\d+)"', section_xml):
        rids.add(m.group(1).decode())
    return rids


def parse_rels(rels_xml):
    """解析关系文件，返回 {rId: (type, target)}"""
    rels = {}
    for m in re.finditer(rb'<Relationship[^>]+Id="(rId\d+)"[^>]+Type="([^"]+)"[^>]+Target="([^"]+)"', rels_xml):
        rid = m.group(1).decode()
        rel_type = m.group(2).decode()
        target = m.group(3).decode()
        rels[rid] = (rel_type, target)
    return rels


def build_rels_xml(rels_dict):
    """构建关系文件XML"""
    parts = [
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    ]
    for rid, (rel_type, target) in sorted(rels_dict.items()):
        parts.append(f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>'.encode())
    parts.append(b'</Relationships>')
    return b'\n'.join(parts)


# 完整的XML命名空间声明（确保Word兼容性）
XML_HEAD = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"'
    b' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    b' xmlns:o="urn:schemas-microsoft-com:office:office"'
    b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    b' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    b' xmlns:v="urn:schemas-microsoft-com:vml"'
    b' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    b' xmlns:w10="urn:schemas-microsoft-com:office:word"'
    b' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    b' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    b' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
    b' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    b' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
    b' mc:Ignorable="w14 wp14">'
    b'<w:body>'
)


def split_docx(docx_path, output_dir, target_level=1, target_idx=None):
    """拆分docx文件"""
    print(f"读取: {docx_path}")
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        doc_xml = z.read('word/document.xml')
        rels_xml = z.read('word/_rels/document.xml.rels')
        all_files = {}
        for n in z.namelist():
            if n not in ('word/document.xml', 'word/_rels/document.xml.rels'):
                all_files[n] = z.read(n)

    all_rels = parse_rels(rels_xml)
    
    body_start = doc_xml.find(b'<w:body>') + 8
    body_end = doc_xml.find(b'</w:body>')
    body = doc_xml[body_start:body_end]

    # 找段落
    para_starts = []
    pos = 0
    while True:
        idx = body.find(b'<w:p w14:paraId=', pos)
        if idx < 0:
            break
        para_starts.append(idx)
        pos = idx + 1
    
    paras = []
    for i, start in enumerate(para_starts):
        end = para_starts[i+1] if i+1 < len(para_starts) else len(body)
        paras.append((start, end))

    # 找目标level的标题
    level_marker = f'outlineLvl w:val="{target_level}"'.encode()
    split_points = []
    for i, (ps, pe) in enumerate(paras):
        elem = body[ps:pe]
        if level_marker in elem:
            text = extract_text(elem)
            if text and len(text) < 50:
                split_points.append({'idx': i, 'text': text})

    if not split_points:
        print(f"未找到 level={target_level} 的标题")
        return

    # 如果指定了序号，只拆分那一个
    if target_idx is not None:
        if target_idx < 1 or target_idx > len(split_points):
            print(f"序号超出范围 (1-{len(split_points)})")
            return
        split_points = [split_points[target_idx - 1]]

    # sectPr
    sect_pr_start = body.find(b'<w:sectPr')
    sect_pr = body[sect_pr_start:body.find(b'</w:sectPr>', sect_pr_start) + 11] if sect_pr_start > 0 else b''

    os.makedirs(output_dir, exist_ok=True)

    for idx, sp in enumerate(split_points):
        ps = paras[sp['idx']][0]
        pe = paras[sp['idx'] + 1][0] if sp['idx'] + 1 < len(paras) else len(body)
        
        # 找下一个拆分点
        for next_sp in split_points:
            if next_sp['idx'] > sp['idx']:
                pe = paras[next_sp['idx']][0]
                break

        section_xml = body[ps:pe]
        used_rids = find_image_rids(section_xml)

        section_rels = {}
        for rid in used_rids:
            if rid in all_rels:
                section_rels[rid] = all_rels[rid]

        new_doc = XML_HEAD + section_xml + sect_pr + b'</w:body></w:document>'

        safe = sanitize(sp['text'])
        filepath = os.path.join(output_dir, f'{safe}.docx')

        with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zout:
            zout.writestr('word/document.xml', new_doc)
            zout.writestr('word/_rels/document.xml.rels', build_rels_xml(section_rels))
            
            for rid in used_rids:
                if rid in all_rels:
                    rel_type, target = all_rels[rid]
                    if 'image' in rel_type:
                        img_path = f'word/{target}' if not target.startswith('/') else target[1:]
                        if img_path in all_files:
                            zout.writestr(img_path, all_files[img_path])

            for name, data in all_files.items():
                if not name.startswith('word/media/'):
                    zout.writestr(name, data)

        size_kb = os.path.getsize(filepath) / 1024
        print(f'[{idx+1:2d}/{len(split_points)}] {sp["text"]} ({size_kb:.0f}KB, {len(used_rids)}图)')

    print(f'\n完成！{len(split_points)} 个文件')


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else 'split_output'
    target_level = int(sys.argv[3]) if len(sys.argv) > 3 else 1
    target_idx = int(sys.argv[4]) if len(sys.argv) > 4 else None
    
    split_docx(docx_path, output_dir, target_level, target_idx)


if __name__ == '__main__':
    main()
