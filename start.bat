@echo off
chcp 65001 >nul
title 文件上传 - 题目生成器

echo ========================================
echo   文件上传 → 题目生成器
echo ========================================
echo.

:: 检查 Python (使用 py 启动器)
py --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 安装依赖
echo [1/2] 正在安装依赖...
py -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --quiet
if %errorlevel% neq 0 (
    echo [警告] 依赖安装可能不完整，继续启动...
)

:: 启动服务
echo.
echo [2/2] 正在启动服务...
echo.
echo 请在浏览器中打开: http://127.0.0.1:5000
echo 按 Ctrl+C 可停止服务
echo ========================================
echo.

py app.py

pause
