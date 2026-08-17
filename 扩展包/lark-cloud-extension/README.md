# Lark Cloud Extension Pack

## 目标

本扩展包用于把飞书云端协同能力从主流程中剥离出来，确保项目默认以“本地优先、离线生成”为主，而飞书能力仅在需要时作为增强扩展启用。

## 适用场景

- 飞书文档读取与同步
- 飞书演示文稿导出与发布
- 知识库/多维表联动
- 云端协同审阅与交付

## 非目标

- 不作为主生成链路前置条件
- 不阻断本地 PPTX / PDF / 其他离线产物输出
- 不要求在无 Lark 环境时被阻塞

## 设计原则

1. 本地优先
   - 默认场景中，项目以本地脚本、模板和输出目录作为正式交付路径。
2. 云端可选
   - 飞书能力由扩展包提供，只有在 `--export lark` 或显式云端协同需求时才调用。
3. 失败不阻断主流程
   - 如果未配置 `lark-cli` 或账号权限，主流程仍继续完成本地生成。

## 目录结构

```text
扩展包/lark-cloud-extension/
├── README.md              # 扩展包说明
├── SKILL.md               # 云端扩展技能说明
├── usage.md               # 使用方式与典型命令
└── references/            # 可选飞书协议/授权/导出参考（按需扩展）
```

## 典型使用方式

### 1. 默认本地生成（推荐）

```bash
cd 生成脚本
node generate.js --theme charcoal-minimal --output demo_local.pptx
```

### 2. 显式启用云端扩展

```bash
cd 生成脚本
node generate.js --theme charcoal-minimal --output demo_cloud.pptx --export lark
```

### 3. 强制本地模式

```bash
cd 生成脚本
node generate.js --local-only --output demo_offline.pptx
```

## 实现说明

主生成器与本地核心能力位于：
- [生成脚本/generate.js](../../生成脚本/generate.js)
- [生成脚本/adapters/lark-export.js](../../生成脚本/adapters/lark-export.js)

本扩展包只提供：
- 适配说明
- 调用方式
- 云端能力启用策略
- 可选的飞书协同规则

## 限制

本扩展包不等于“完全移除 Lark 能力”，而是将其收敛为可选扩展层。若项目需要深度接入飞书文档、知识库或演示文稿平台，仍应在此扩展包内补充更完整的适配逻辑。

## 结论

本扩展包是为实现“本地优先、云端可选”的架构而设计，确保主程序不再依赖飞书云端环境，项目可在离线条件下完成主体交付。
