@echo off
chcp 65001 >nul 2>nul
setlocal
cd /d "%~dp0"
title Bank Training Flowchart Tool

where py >nul 2>nul
if not errorlevel 1 (
    start "" /b py -3 "%~dp0flowchart_tool_launcher.pyw"
    exit /b 0
)

where python >nul 2>nul
if not errorlevel 1 (
    start "" /b python "%~dp0flowchart_tool_launcher.pyw"
    exit /b 0
)

rem ASCII-only fallback messages to avoid codepage issues
echo [ERROR] Python 3 not found. Please install Python 3 first.
echo Download: https://www.python.org/downloads/
pause
exit /b 1
