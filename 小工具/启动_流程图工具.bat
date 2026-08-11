@echo off
chcp 65001 >nul
title 银行培训流程图离线工具 - 启动器
echo ================================================
echo   银行培训 · Word拆分 / PPT流程图 离线工具
echo ================================================
echo.
echo 正在启动图形界面工具，请稍候...
echo （若长时间无窗口弹出，请确认已安装 Python 3）
echo.

where py >nul 2>nul
if %errorlevel% equ 0 (
    start "" py -3 "%~dp0流程图_小工具.py"
) else (
    rem 尝试 python 命令
    where python >nul 2>nul
    if %errorlevel% equ 0 (
        start "" python "%~dp0流程图_小工具.py"
    ) else (
        echo [错误] 未检测到 Python，请先安装 Python 3 运行环境。
        echo 下载地址: https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

exit /b 0
