# quiz-extraction-skill 技能索引

**版本**: v1.0.0
**发布日期**: 2026-08-18

## 技能概述

从 Word/PPT/TXT/Markdown 源文档抽取考试题，AI 语义识别题型与题组，脚本做解析与规范化，输出多格式题库。

## 功能列表

- 源文档解析（docx/pptx/txt/md）为统一中间结构
- AI 语义抽取题组 JSON（对齐 quiz_schema）
- 后端规范化（清洗/合并/组号重排/答案归一）
- schema 验收门禁（validate_quiz.py）
- 多格式输出（JSON/CSV/Markdown/PPTX）

## 使用方式

按 `SKILL.md` 五步工作流执行，核心命令：

```bash
python skills/quiz-extraction-skill/tools/parse_quiz_source.py 输入.docx -o 输出/中间结构.json
python skills/quiz-extraction-skill/tools/normalize_quiz.py 输出/AI抽取.json -o 输出/规范化.json
python skills/quiz-extraction-skill/tools/validate_quiz.py 输出/规范化.json
python skills/quiz-extraction-skill/tools/export_quiz.py all 输出/规范化.json -o 输出/题库
```

## 目录结构

- `SKILL.md` — 主技能
- `schemas/` — quiz_schema.json 统一 schema
- `templates/` — CSV 模板
- `presets/` — 题型配色预设
- `styles/` — 样式示例
- `examples/` — 示例输入与产物
- `domain/` — 题型/题组/输出格式规则
- `tools/` — 解析/规范化/校验/输出脚本

## 相关技能

- 无直接依赖；可与 ppt-framework（PPT 生成）组合使用

## 文档

- SKILL_INDEX.md
- README.md
- SKILL.md