---
name: "docx-splitter-skill"
description: "Word文档章节拆分技能：根据标题层级智能拆分docx文件，支持自定义层级、生成目录索引、保持格式完整。Invoke when splitting Word documents by heading levels."
---

# DocxSplitterSkill Word文档章节拆分技能

> 版权声明详见 `../../shared/references/COPYRIGHT.md`

## 1. 触发规则

- **触发时机**：用户要求拆分Word文档、分析文档结构、提取特定章节
- **入参**：`{"action": "analyze_structure / split_document / batch_split", "input_file": "docx路径", "level": "拆分层级", "output_dir": "输出目录"}`

## 2. 流程

### 环节 1：文档结构分析（analyze_structure）

**执行内容**：
1. 解压docx文件获取XML
2. 解析document.xml提取段落
3. 识别标题样式和大纲级别
4. 构建层级树结构

**输出**：
```
# 文档标题
  ## 章节1
    ### 子章节1.1
    ### 子章节1.2
  ## 章节2
    ### 子章节2.1
```

### 环节 2：章节拆分（split_document）

**执行内容**：
1. 按指定层级切分段落
2. 为每个章节生成独立docx
3. 保留原格式（样式、字体）
4. 生成章节索引

**拆分规则**：
| 层级 | 切分点 | 包含内容 |
|------|--------|----------|
| Level 1 | 标题1 | 该标题及其下所有内容 |
| Level 2 | 标题2 | 该标题及其下所有内容直到下一个同级标题 |
| Level N | 标题N | 同上 |

### 环节 3：批量拆分（batch_split）

**执行内容**：
1. 扫描目录下所有docx文件
2. 按统一规则拆分
3. 生成汇总报告

## 3. 输出规范

| 文件 | 说明 |
|------|------|
| `section_NNN_标题.docx` | 拆分后的章节文件 |
| `INDEX.md` | Markdown格式索引 |
| `split_report.json` | JSON格式结构报告 |

### split_report.json 结构

```json
{
  "total_sections": 4,
  "sections": [
    {
      "index": 1,
      "heading": "章节标题",
      "level": 2,
      "parent_heading": "父级标题",
      "paragraph_count": 15,
      "filename": "section_001_章节标题.docx"
    }
  ]
}
```

## 4. 命令行接口

```bash
# 查看文档结构
py -3 tools/docx_splitter.py --structure input.docx

# 按层级拆分
py -3 tools/docx_splitter.py --split input.docx --level 2

# 指定输出目录
py -3 tools/docx_splitter.py --split input.docx --level 2 --output ./split

# 生成报告
py -3 tools/docx_splitter.py --split input.docx --level 2 --report

# JSON输出
py -3 tools/docx_splitter.py --structure input.docx --json
```

## 5. 行业最佳实践

### 5.1 拆分策略选择

| 文档类型 | 推荐层级 | 理由 |
|----------|---------|------|
| 操作手册 | Level 3 | 按操作步骤拆分，便于培训 |
| 制度文档 | Level 2 | 按条款拆分，便于检索 |
| 培训PPT素材 | Level 1 | 按大章节拆分，便于制作 |
| 技术规范 | Level 2 | 按功能模块拆分 |

### 5.2 格式保持

- 保留段落样式（标题、正文、列表）
- 保留字体和字号设置
- 保留表格结构
- 保留图片引用

### 5.3 质量检查

拆分后检查：
- [ ] 章节内容完整
- [ ] 格式无丢失
- [ ] 图片正常显示
- [ ] 表格结构正确

## 6. 边界（刹车规则）

- 仅支持.docx格式（不支持旧版.doc）
- 文档必须使用标准标题样式
- 复杂交叉引用可能失效
- 嵌入式对象需要手动处理

---

**文档版本**：v1.0.0　**最后更新**：2026-08-07
**知识产权所有**：段波（验证邮箱：duanbo.douglas@163.com）
