---
name: "document-processing"
description: "文档处理技能包：Word文档章节拆分、结构分析、格式转换。触发词：拆分文档、章节拆分、docx拆分、文档结构、文档转换。Load when the user wants to split Word documents by heading levels, analyze document structure, or convert document formats."
---

# DocumentProcessingSkill 文档处理技能包

> 版权：`../references/COPYRIGHT.md`　Token：`../references/token_standard.md`

## 1. 元数据

- **技能名称**：document-processing
- **技能版本**：v1.1.0
- **发布日期**：2026-08-07
- **参考标准**：OOXML (ISO/IEC 29500) · IEEE 830 · BABOK v3

## 2. 触发规则

用户表达「拆分文档/章节拆分/docx拆分/文档结构/文档转换/Word拆分」时加载本包。

## 3. 流程（路由到 domain/）

| 环节 | action | 明细 |
|------|--------|------|
| 文档结构分析 | analyze_structure | `domain/docx-splitter.md` |
| 章节拆分 | split_document | `domain/docx-splitter.md` |
| 批量拆分 | batch_split | `domain/docx-splitter.md` |
| 格式转换 | convert_format | `domain/docx-converter.md` |

## 4. 核心能力

### 4.1 Word文档章节拆分

根据Word文档的大纲级别（outlineLvl）进行智能拆分：

```bash
# 拆分所有level=1的章节（默认）
py -3 tools/split_docx_by_level.py input.docx output_dir

# 只拆分第3个level=1章节
py -3 tools/split_docx_by_level.py input.docx output_dir 1 3

# 按level=2拆分
py -3 tools/split_docx_by_level.py input.docx output_dir 2
```

### 4.2 行业最佳实践

| 实践 | 说明 |
|------|------|
| 只复制引用的图片 | 扫描`<a:blip>`标签提取rId，只复制被引用的图片，大幅减小文件体积 |
| 完整命名空间声明 | 包含所有必要的XML命名空间（w, r, wp, wps等），确保Word兼容性 |
| 流式处理 | 不解析完整XML树，支持超大文档（10MB+） |
| 保留格式 | 拆分后文档保留原格式（样式、字体、图片） |
| 关系文件同步 | 只包含被引用的关系，避免冗余 |

### 4.3 拆分策略

| 文档类型 | 推荐拆分层级 | 说明 |
|----------|-------------|------|
| 操作手册 | Level 3 | 按"步骤"拆分 |
| 制度文档 | Level 2 | 按"条款"拆分 |
| 培训材料 | Level 1 | 按"章节"拆分 |
| 技术规范 | Level 2 | 按"章节"拆分 |

## 5. 技术实现

### 5.1 大纲级别识别

Word文档使用`outlineLvl`属性标识标题层级：
- `outlineLvl w:val="0"` → 一级标题
- `outlineLvl w:val="1"` → 二级标题
- `outlineLvl w:val="2"` → 三级标题
- 以此类推...

### 5.2 图片引用机制

Word文档中的图片通过关系ID引用：
1. XML中：`<a:blip r:embed="rId10"/>`
2. 关系文件：`<Relationship Id="rId10" Target="media/image1.png"/>`
3. 图片文件：`word/media/image1.png`

### 5.3 文件大小优化

| 方案 | 原始文件 | 拆分后 | 说明 |
|------|----------|--------|------|
| 复制全部资源 | 356MB | 300MB+ | 包含所有图片 |
| 只复制引用资源 | 356MB | 3-24MB | 只包含被引用的图片 |

## 6. 输出规范

- 拆分文件：`标题.docx`
- 文件大小：通常3-24MB（取决于图片数量）

## 7. 边界

- 仅处理.docx格式（不支持.doc）
- 需要文档使用标准大纲级别（outlineLvl）
- 复杂表格和嵌入对象可能需要手动调整

---

**文档版本**：v1.1.0　**最后更新**：2026-08-07
**知识产权所有**：段波（验证邮箱：duanbo.douglas@163.com）
