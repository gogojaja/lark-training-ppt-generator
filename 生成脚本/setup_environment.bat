# 生成脚本运行环境准备（setup_environment.bat）
@echo off
chcp 65001 >nul
echo ============================================
echo   lark-workflow-handbook-deck 生成脚本环境准备
echo ============================================
echo.

rem ---- 1. 检查 node ----
where node >nul 2>nul
if %errorlevel%==0 (
  echo [OK] node 已安装: 
  node --version
) else (
  echo [X] 未检测到 node。生成脚本 generate.js 依赖 node + pptxgenjs。
  echo     请先安装 Node.js LTS（官网 https://nodejs.org 下载安装包），
  echo     安装后重新运行本脚本。
  echo.
  echo     备选：本机可用 py -3 运行 Python，如需纯 Python 生成方案请联系维护者。
  goto :end
)

rem ---- 2. 安装 pptxgenjs ----
if exist node_modules (
  echo [OK] 已存在 node_modules，跳过依赖安装
) else (
  echo [..] 正在安装 pptxgenjs ...
  npm install pptxgenjs
  if %errorlevel% neq 0 (
    echo [X] 依赖安装失败，请检查网络后重试。
    goto :end
  )
  echo [OK] 依赖安装完成
)

echo.
echo [OK] 环境就绪。运行方式：
echo   1. 生成主 PPT： node generate.js        （输出 综合个人开户_柜面操作培训.pptx）
echo   2. 生成风格样例：node scripts/generate_style_samples.js
echo.
echo 提示：生成前确认素材目录 images/ 存在背景图（bg-cover_16x9.jpg / bg-closing_16x9.jpg）。

:end
echo.
echo 环境准备结束。
pause
