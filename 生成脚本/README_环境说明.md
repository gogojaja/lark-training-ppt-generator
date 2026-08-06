# 生成脚本运行环境说明

「制度手册转宣讲 PPT 工作流」的本地生成脚本基于 **Node.js + pptxgenjs** 实现，可离线生成 PPTX 样例（不依赖 lark-cli）。

> **零依赖流程图工具**：`gen_flowchart.py` 用 Python 标准库直接读写 PPTX/docx（zip+xml），**无需 node/pip/第三方库**，本机 `py -3` 即可运行。

## 环境要求

| 组件 | 说明 |
|------|------|
| Node.js | LTS 版（未安装时运行 `setup_environment.bat` 会提示） |
| pptxgenjs | 由 `setup_environment.bat` 自动 `npm install` |

## 快速开始

```bat
:: 1. 一键准备环境（检查 node + 安装 pptxgenjs）
setup_environment.bat

:: 2. 生成主 PPT（读取 slide_plan.json 对应的 18 页版式）
node generate.js

:: 3. 生成风格样例（Step 3 风格预览用）
node scripts/generate_style_samples.js
```

## Word → 一页流程图 PPT（零依赖）

```bat
:: 方式一：读 Word 文档（第一段作标题，其余段落按顺序作为流程步骤）
py -3 gen_flowchart.py 流程操作.docx --out 流程图.pptx

:: 方式二：命令行直接给步骤（分号或换行分隔）
py -3 gen_flowchart.py --steps "预约叫号;身份识别;资料录入;业务开通;授权复核;归档结束" --title 柜面开户操作流程 --out 流程图.pptx
```

- 输出为**一页 16:9 纵向流程图**：开始（红）→ 步骤（蓝）→ 结束（绿），方框 + 下箭头。
- 不依赖 node/python-pptx/pip，仅用 Python 标准库（zipfile+ElementTree 手写 PPTX XML）。

## 输出

- `generate.js` → `综合个人开户_柜面操作培训.pptx`
- `generate_style_samples.js` → `style_samples.pptx`（风格变体预览）
- `gen_flowchart.py` → 指定 `--out` 的一页流程图 PPTX

## 与本机现状

> **注意**：当前执行机未安装 node。可直接运行 `setup_environment.bat` 获得安装指引；
> 也可用 `py -3`（已可用）运行 `D:/trae/lark-slides/scripts/xml_text_overlap_check.py` 做回读验证。

## 依赖外部技能

工作流在飞书侧生成演示文稿依赖 `lark-cli` 与 5 个 lark-* 技能（见项目 SKILL.md 引用）。

---

**文档版本**：v1.0.0　**最后更新**：2026-08-06