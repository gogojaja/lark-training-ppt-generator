# Skill: quiz-extraction-skill

# 考试题抽取技能包

> 版权：`../shared/references/COPYRIGHT.md`

## 1. 元数据

- **技能名称**：quiz-extraction-skill
- **技能版本**：v1.0.0
- **发布日期**：2026-08-18
- **技能定位**：从 Word/PPT/TXT/Markdown 源文档中抽取考试题，AI 语义识别题型与题组，脚本做解析与规范化，输出 JSON/CSV/Markdown/PPTX 多格式题库
- **适用场景**：培训考核出题、试卷数字化入库、题库建设、课件考点提炼
- **依赖库**：python-docx≥1.0、python-pptx≥1.0（均已安装，零额外安装）

## 2. 五步工作流

```
源文档 → ①文档解析 → ②AI语义抽取 → ③后端规范化 → ④验收门禁 → ⑤多格式输出
```

### Step 1：文档解析（脚本，机械层）

**工具**：`tools/parse_quiz_source.py`（已内置）

```bash
python tools/parse_quiz_source.py 输入文档/试卷.docx -o 输出/中间结构.json
python tools/parse_quiz_source.py 输入文档/课件.pptx -o 输出/中间结构.json
python tools/parse_quiz_source.py 输入文档/试题.md   -o 输出/中间结构.json
```

**最佳实践**：
- Word 优先按段落与表格结构解析；PPT 按幻灯片顺序解析，每页聚合为独立文本块
- 大文档（>50MB）先拆分章节再逐章抽取，避免上下文溢出
- 解析后检查 `blocks` 数量与内容，确认文本完整无乱码

### Step 2：AI 语义抽取（模型层，核心）

**架构依据**：对齐行业最佳实践「文档解析打底 + 大模型语义抽取 + 后端稳定化」。AI 只负责语义判断，机械排版/清洗/校验由脚本固化，遵循项目「**模型无关 / 工具先于模型**」原则。

**工作流**：

```
┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐
│ ①读取中间结构 │ → │ ②识别题组/题型 │ → │ ③判断题干选项答案 │ → │ ④输出题组JSON │
└──────────────┘  └──────────────┘  └──────────────┘  └──────────────┘
```

**AI 语义抽取核心 Prompt 模板**（skill 执行者将中间结构 JSON 与以下规则交给模型）：

> 请阅读以下文档中间结构（JSON），抽取其中的考试题，输出符合 quiz_schema.json 的题组 JSON。严格遵循：
>
> **一、题型识别**
> 1. `single_choice` 单选：题干 + 选项，答案 1 项
> 2. `multiple_choice` 多选/不定项：答案可多项
> 3. `true_false` 判断：题干陈述，答案「正确/错误」
> 4. `fill_blank` 填空：题干含下划线/括号空位，答案对应每空
> 5. `short_answer` 简答：参考答案要点
>
> **二、题组结构**
> 1. 阅读材料/案例对应多道小题时，用共享题干 `shared_stem` 组成题组（group），小题挂 `questions`，不得重复抄写材料
> 2. 题组 `section` 填板块标题（如「一、单选题」）；`source_pages` 填涉及页号
>
> **三、字段规则**
> 1. `stem` 题干完整；选项用 `options` 数组（如 `["A. xxx","B. xxx"]`），答案 `answer` 填选项字母（判断题填「正确/错误」）
> 2. `analysis` 尽量给出解析；`difficulty` 按 `easy/medium/hard` 分级
> 3. 识别不出答案的题目，`answer` 留空数组，不要臆造
>
> **四、输出要求**
> 1. 输出合法 JSON，不要包裹 Markdown 代码块说明
> 2. 必须包含顶层 `paper_title`（试卷标题）与 `groups` 数组
> 3. 原文档中明确标注的「参考答案」「正确答案」等必须忠实抽取，不得改写

### Step 3：后端规范化（脚本，收敛层）

**工具**：`tools/normalize_quiz.py`

```bash
python tools/normalize_quiz.py 输出/AI抽取.json -o 输出/规范化.json
```

**处理动作**：
- JSON 兜底解析（容忍 AI 输出包裹代码块/首尾空白）
- 字段类型清洗：score 转数字、source_page 转整数、空字符串归一
- 组号重排（G1,G2,…）与小题 qid 重排（G1-Q1,…）
- 答案归一：大写、去分隔符、数组化；判断题答案映射为「正确/错误」
- 自动合并共享题干题组

### Step 4：验收门禁

**工具**：`tools/validate_quiz.py`

```bash
python tools/validate_quiz.py 输出/规范化.json        # 常规校验
python tools/validate_quiz.py 输出/规范化.json --strict  # 含警告也视为失败
```

**通过条件**：
- 顶层含 `paper_title` / `source` / `groups`
- 每道题 `qid` / `type`（五类枚举） / `stem` 非空，`qid` 全局唯一
- 选择题选项非空且答案落在选项内；判断题答案限「正确/错误」
- 门禁不通过返回非 0 退出码并列出错误/警告清单

**处理**：门禁未过 → 回 Step 2 让 AI 按错误清单修订，或人工校正后重新规范化。

### Step 5：多格式输出

**工具**：`tools/export_quiz.py`

```bash
python tools/export_quiz.py csv  输出/规范化.json -o 输出/题库.csv      # UTF-8 BOM，Excel 直开
python tools/export_quiz.py md   输出/规范化.json -o 输出/试卷.md
python tools/export_quiz.py pptx 输出/规范化.json -o 输出/问卷.pptx
python tools/export_quiz.py all  输出/规范化.json -o 输出/目录           # 一次全出
```

**CSV 列**：`group_id,section,shared_stem,qid,type,stem,option_a~option_f,answer,analysis,difficulty,score,source_page`

## 3. 目录结构

```
skills/quiz-extraction-skill/
├── SKILL.md                  # 本文件
├── SKILL_INDEX.md            # 技能索引
├── README.md                 # 使用说明
├── schemas/quiz_schema.json  # 统一题库 schema
├── templates/quiz_template.csv  # CSV 模板
├── presets/                  # 题型配色预设
├── styles/                   # 样式示例
├── examples/                 # 示例（输入/中间/规范化/输出）
├── domain/                   # 领域规则（题型/题组/格式规范）
└── tools/                    # 解析/规范化/校验/输出脚本
```

## 4. 题型与输出格式规范

- 题型枚举与识别规则见 `domain/question-types.md`
- 题组结构（共享题干/材料题）见 `domain/group-structure.md`
- 统一 schema 见 `schemas/quiz_schema.json`（QTI 3.0 思想裁剪）
- CSV/JSON/Markdown/PPTX 输出格式规范见 `domain/output-format.md`

---

## 闭环执行系统

> 适用范围：quiz-extraction-skill 的执行与验收。结构对齐 skill-maintenance 通用模板。

### 1. 任务入口

- **输入**：源文档（.docx/.pptx/.txt/.md）；触发场景：培训考核出题、试卷数字化、题库建设、课件考点提炼
- **前置**：源文件可读；python-docx/python-pptx 已安装；工作目录可写；用户指定输出格式（缺省 CSV）
- **不适用**：源文档为扫描件/图片（需 OCR，不在本技能范围）；源文档无题目结构（应转「出题生成」场景而非抽取）

### 2. 执行状态

| 状态 | 进入条件 | 退出条件 | 处理方式 |
|------|---------|---------|---------|
| 待启动 | 任务明确且前置满足 | 用户确认 | 准备源文件与输出目录 |
| 执行中 | 任务启动 | Step 1~3 完成 | 按五步工作流执行 |
| 校验中 | 规范化完成 | 门禁通过/失败 | 运行 validate_quiz.py |
| 阻塞 | 解析失败/依赖缺失 | 补充信息/人工处理 | 暂停并记录原因 |
| 完成 | 门禁通过且输出生成 | 进入交接 | 归档证据与产出 |
| 回退 | 门禁未过/输出异常 | 回到最近稳定状态 | 修订 AI 结果或重跑 |

### 3. 执行动作层

- 执行步骤 1：`parse_quiz_source.py` 解析源文档为中间结构 JSON
- 执行步骤 2：AI 按核心 Prompt 语义抽取题组 JSON（复用 `schemas/quiz_schema.json`）
- 执行步骤 3：`normalize_quiz.py` 规范化/清洗/合并/重排
- 执行步骤 4：`validate_quiz.py` 验收门禁
- 执行步骤 5：`export_quiz.py` 多格式输出
- 所需工具/脚本：`tools/` 下 4 个脚本；领域规则 `domain/`
- 输入输出约束：中间产物与最终产物落盘 `输出/` 目录；统一 UTF-8（CSV 为 UTF-8 BOM）

### 4. 验收门禁

- 必须产出物：规范化 JSON；至少一种输出（CSV/Markdown/PPTX）；门禁通过记录
- 通过条件：`validate_quiz.py` 无错误退出；抽取题目与源文档语义一致；选择题答案落在选项内
- 失败条件：`qid` 重复、题型非法、题干为空、选择题答案越界、判断题答案非法
- 审核对象：用户或 skill 执行者复核源文档对照

### 5. 失败处理

- 失败类型：格式不支持、JSON 解析失败、门禁未过、依赖缺失
- 恢复策略：格式不支持→先转 .docx/.pptx/.md；JSON 失败→回 Step 2 重出或手工修
- 回滚方案：保留 AI 抽取原始输出（`输出/AI抽取.json`），规范化失败可回滚重跑
- 重试策略：门禁错误清单反馈 AI 修订，最多重试 3 次，仍失败转人工
- 是否需要人工确认：判断题答案映射、填空题多空答案需要人工复核时确认

### 6. 产出与交接

- 产出物列表：中间结构 JSON、AI 抽取 JSON、规范化 JSON、CSV/Markdown/PPTX 题库
- 保存路径：`输出/` 目录
- 交接对象：题库使用者/培训系统（CSV 入库、PPTX 问卷）
- 下一步动作：题库校对 → 发布培训考核
- 归档条件：门禁通过、输出验证可打开、源文档路径留痕

### 7. 审计记录

- 执行时间：记录开始/结束时间
- 关键参数：源文档路径、题型、输出格式、模型版本
- 关键决策：是否通过门禁、是否回退、是否人工确认
- 结果证据：`输出/` 目录产物与校验日志
- 失败原因：留痕于 `输出/校验错误.log`