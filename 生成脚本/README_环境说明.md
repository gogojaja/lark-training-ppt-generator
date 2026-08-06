# 生成脚本运行环境说明

「制度手册转宣讲 PPT 工作流」的本地生成脚本基于 **Node.js + pptxgenjs** 实现，可离线生成 PPTX 样例（不依赖 lark-cli）。

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

## 输出

- `generate.js` → `综合个人开户_柜面操作培训.pptx`
- `generate_style_samples.js` → `style_samples.pptx`（风格变体预览）

## 与本机现状

> **注意**：当前执行机未安装 node。可直接运行 `setup_environment.bat` 获得安装指引；
> 也可用 `py -3`（已可用）运行 `D:/trae/lark-slides/scripts/xml_text_overlap_check.py` 做回读验证。

## 依赖外部技能

工作流在飞书侧生成演示文稿依赖 `lark-cli` 与 5 个 lark-* 技能（见项目 SKILL.md 引用）。

---

**文档版本**：v1.0.0　**最后更新**：2026-08-06