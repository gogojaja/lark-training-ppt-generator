# 输出格式规范

> quiz-extraction-skill 领域规则：四种输出形态（JSON/CSV/Markdown/PPTX）的字段与编码约定。

## 1. JSON（中间产物与主交付物）

- **文件**：`输出/规范化.json`（主）、`输出/AI抽取.json`（AI 原始）
- **编码**：UTF-8（无 BOM），`ensure_ascii=False`，缩进 2 空格
- **结构**：对齐 `schemas/quiz_schema.json`，顶层 `paper_title / source / metadata / groups`

## 2. CSV（Excel 直接可打开）

- **文件**：`输出/题库.csv`
- **编码**：UTF-8 **BOM**（Excel 双击不乱码）
- **列**：`group_id,section,shared_stem,qid,type,stem,option_a,option_b,option_c,option_d,option_e,option_f,answer,analysis,difficulty,score,source_page`
- 选项映射：`option_a~option_f` 按顺序映射前 6 个选项，多余选项留空并记录 warning
- `answer` 多项用 `/` 连接；判断题归一为「正确/错误」；填空题多项用 `/` 对应每空

## 3. Markdown（人工阅读/打印试卷）

- **文件**：`输出/试卷.md`
- 结构：`# 试卷标题` → 每板块 `## 板块标题` → 材料题块 → 每题 `**qid（题型）题干**` → `- 选项` → `> 答案` / `> 解析`
- 编码：UTF-8（无 BOM）

## 4. PPTX（问卷/课堂展示）

- **文件**：`输出/问卷.pptx`
- 布局：每个板块一页（标题=板块，正文=材料+题目），题目字号 16pt、选项 14pt
- 依据：python-pptx 生成，兼容 Office/PowerPoint/WPS
- 注：PPTX 为展示用途，入库仍以 CSV/JSON 为准

## 5. 通用约定

| 项 | 约定 |
|----|------|
| 字体 | 默认（宋体/系统字体），不内嵌字体 |
| 图片/公式 | 源文档中图片以路径/页号定位，不做 OCR |
| 空值 | 空字符串统一为 `""`；未知数字为 `null`（CSV 中为空） |
| 文件名 | `输出/` 目录统一落盘，文件名含中文 |
| 版本 | schema 变更需同步 `schemas/quiz_schema.json` 与四个 export 工具 |