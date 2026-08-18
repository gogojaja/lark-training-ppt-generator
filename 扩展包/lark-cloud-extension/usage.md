# Lark Cloud Extension Usage

## 1. 默认使用

默认场景下，项目无需安装飞书 CLI，也无需登录飞书账号即可生成 PPTX：

```bash
cd 生成脚本
node generate.js --theme charcoal-minimal --output demo_local.pptx
```

## 2. 显式启用 Lark 扩展

如果需要把本地输出同步到飞书，可使用：

```bash
cd 生成脚本
node generate.js --theme charcoal-minimal --output demo_cloud.pptx --export lark
```

## 3. 强制关闭云端能力

```bash
cd 生成脚本
node generate.js --local-only --output demo_offline.pptx
```

## 4. 失败处理

若 `lark-cli` 未安装或未授权：
- 本地 PPTX 仍会正常生成
- 扩展输出会返回提示信息
- 主流程不会失败

## 5. 设计意图

本扩展包的核心价值不是替代本地生成，而是：
- 在需要时提供云端协同能力
- 在不需要时不增加额外门槛
- 在云端能力不可用时仍保持项目正常交付
