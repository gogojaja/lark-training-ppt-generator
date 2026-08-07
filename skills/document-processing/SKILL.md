---
name: "document-processing"
description: "文档处理技能包：Word文档章节拆分、结构分析、格式转换。触发词：拆分文档、章节拆分、docx拆分、文档结构、文档转换。Load when the user wants to split Word documents by heading levels, analyze document structure, or convert document formats."
---

# DocumentProcessingSkill 文档处理技能包

> 版权：`../references/COPYRIGHT.md`　Token：`../references/token_standard.md`

## 1. 元数据

- **技能版本**：v1.0.0
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

根据Word文档的标题层级（Heading 1/2/3...）进行智能拆分：

```bash
# 查看文档结构
py -3 tools/docx_splitter.py --structure input.docx

# 按指定层级拆分
py -3 tools/docx_splitter.py --split input.docx --level 2

# 拆分并生成报告
py -3 tools/docx_splitter.py --split input.docx --level 2 --report
```

### 4.2 行业最佳实践

| 实践 | 说明 |
|------|------|
| 保留格式 | 拆分后文档保留原格式（样式、字体、图片） |
| 自定义层级 | 支持按任意标题层级拆分 |
| 目录索引 | 自动生成INDEX.md索引文件 |
| 结构报告 | 输出JSON格式的结构化元数据 |
| 批量处理 | 支持批量拆分多个文档 |

### 4.3 拆分策略

| 文档类型 | 推荐拆分层级 | 说明 |
|----------|-------------|------|
| 操作手册 | Level 3 | 按"步骤"拆分 |
| 制度文档 | Level 2 | 按"条款"拆分 |
| 培训材料 | Level 1 | 按"章节"拆分 |
| 技术规范 | Level 2 | 按"章节"拆分 |

## 5. 输出规范

- 拆分文件：`section_NNN_标题.docx`
- 索引文件：`INDEX.md`
- 结构报告：`split_report.json`

## 6. 边界

- 仅处理.docx格式（不支持.doc）
- 需要文档使用标准标题样式
- 复杂表格和图片可能需要手动调整

---

**文档版本**：v1.0.0　**最后更新**：2026-08-07
**知识产权所有**：段波（验证邮箱：duanbo.douglas@163.com）
