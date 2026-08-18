---
name: lark-cloud-extension
version: 1.0.0
description: "飞书云端协同扩展包：仅在需要时启用，提供文档同步、导出、发布和知识库协同能力，不作为主生成链路前置条件。"
---

# Lark Cloud Extension

## 目标

本技能包用于为本地优先的 PPT 生成项目补充可选云端能力。其设计目标是：
- 让主流程无需依赖飞书即可运行
- 让云端能力作为增强功能存在
- 让云端失败不会中断本地交付

## 适用场景

- 飞书文档读取与内容同步
- 飞书演示文稿导出
- 知识库发布
- 多维表/表格/协同审阅联动

## 模式

### 模式 A：本地优先（默认）

主流程仅依赖本地脚本、模板与生成器，不依赖飞书平台。

### 模式 B：云端增强（可选）

仅在明确需要云端协同时使用：

```bash
cd 生成脚本
node generate.js --output demo_cloud.pptx --export lark
```

如果未安装 `lark-cli`，系统应记录提示并继续本地输出，不应中断主流程。

## 必备条件

- Node.js
- Python（如需流程图/脚本工具）
- 仅在云端增强模式下：`lark-cli` + 飞书账号授权

## 约束

1. 不能成为主流程前置条件
2. 不能阻断离线生成
3. 失败时必须返回提示，不得导致主流程失效
4. 云端能力必须落在扩展层，而非核心生成器中

## 典型调用

```bash
cd 生成脚本
node generate.js --local-only --output demo_offline.pptx
node generate.js --output demo_cloud.pptx --export lark
```

## 说明

主生成逻辑位于：
- [生成脚本/generate.js](../../生成脚本/generate.js)
- [生成脚本/adapters/lark-export.js](../../生成脚本/adapters/lark-export.js)

本扩展包仅负责云端适配、能力说明和增强调用；它不是主程序依赖。
