# quiz-extraction-skill 考试题抽取技能包

从 Word/PPT/TXT/Markdown 源文档中抽取考试题，AI 语义识别题型与题组，脚本做解析与规范化，输出 JSON/CSV/Markdown/PPTX 多格式题库。

## 快速开始

```bash
# 1. 解析源文档为中间结构
python skills/quiz-extraction-skill/tools/parse_quiz_source.py 输入文档/试卷.docx -o 输出/中间结构.json

# 2. AI 语义抽取（按 SKILL.md Step 2 的 Prompt 模板，输出 输出/AI抽取.json）

# 3. 规范化
python skills/quiz-extraction-skill/tools/normalize_quiz.py 输出/AI抽取.json -o 输出/规范化.json

# 4. 验收门禁
python skills/quiz-extraction-skill/tools/validate_quiz.py 输出/规范化.json

# 5. 输出
python skills/quiz-extraction-skill/tools/export_quiz.py all 输出/规范化.json -o 输出/题库
```

## 支持范围

- 输入：`.docx` / `.pptx` / `.txt` / `.md`
- 题型：单选题、多选题、判断题、填空题、简答题、材料题（共享题干题组）
- 输出：JSON、CSV（UTF-8 BOM）、Markdown、PPTX

## 设计原则

- **模型无关 / 工具先于模型**：语义判断（题型/题组/题干）由模型负责，解析/清洗/校验/排版由脚本固化
- **混合架构**：文档解析打底 → AI 语义抽取 → 后端规范化收敛，对齐行业最佳实践
- **可验收**：`validate_quiz.py` 门禁，错误清单可回传 AI 修订

## 目录

- `SKILL.md` — 主技能（五步工作流 + 闭环执行系统）
- `schemas/quiz_schema.json` — 统一题库 schema
- `domain/` — 题型/题组/输出格式规范
- `tools/` — 解析/规范化/校验/输出脚本
- `templates/quiz_template.csv` — CSV 模板
- `examples/` — 示例输入与产物