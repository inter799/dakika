@echo off
chcp 65001 >nul
title 文件上传 - 题目生成器 v2.0

echo ================================================
echo   文件上传 - 题目生成器 v2.0 ^(Windows^)
echo ================================================
echo.

:: ==================== 检测 Python ====================
py --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时务必勾选 "Add Python to PATH"
    pause
    exit /b 1
)

:: 显示 Python 版本
for /f "tokens=*" %%i in ('py --version') do set PY_VER=%%i
echo [检查] %PY_VER% 已就绪
echo.

:: ==================== 安装/更新依赖 ====================
echo [1/3] 正在安装/更新 Python 依赖...
py -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --quiet --upgrade
if %errorlevel% neq 0 (
    echo [警告] 部分依赖安装失败，尝试用官方源重试...
    py -m pip install -r requirements.txt --quiet
)
echo [完成] 依赖安装完成
echo.

:: ==================== 检测关键依赖 ====================
echo [2/3] 检测关键功能组件...

:: 检测 pywin32（.doc 旧版文件支持）
py -c "import win32com.client" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] pywin32 未安装，旧版 .doc 文件解析可能受限
    echo       安装命令: py -m pip install pywin32
) else (
    echo [  OK  ] pywin32 已就绪 - 支持 .doc 旧版文件解析
)

:: 检测 Pillow
py -c "from PIL import Image" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] Pillow 未安装，图片处理功能受限
    echo       安装命令: py -m pip install Pillow
) else (
    echo [  OK  ] Pillow 已就绪 - 支持图片上传
)

:: 检测 Tesseract OCR
py -c "import pytesseract" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] pytesseract 未安装
    echo       安装命令: py -m pip install pytesseract
) else (
    tesseract --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [提示] Tesseract OCR 引擎未安装，图片文字识别不可用
        echo       下载地址: https://github.com/UB-Mannheim/tesseract/wiki
        echo       安装时勾选中文语言包 Chinese ^(Simplified^)
    ) else (
        echo [  OK  ] Tesseract OCR 已就绪 - 支持图片文字识别
    )
)

:: 检测 python-docx
py -c "from docx import Document" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] python-docx 未安装，.docx 文件解析不可用
) else (
    echo [  OK  ] python-docx 已就绪 - 支持 .docx 文件解析
)

:: 检测 PDF 解析
py -c "import pdfplumber" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] pdfplumber 未安装，PDF 解析可能降级
) else (
    echo [  OK  ] PDF 解析库已就绪
)

:: 检测 Excel/PPT 解析
py -c "import openpyxl; from pptx import Presentation" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] openpyxl/pptx 未安装，Excel/PPT 解析不可用
) else (
    echo [  OK  ] Excel/PPT 解析库已就绪
)

:: 检测 AI SDK
py -c "from openai import OpenAI" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] OpenAI SDK 未安装，AI 模式不可用
) else (
    echo [  OK  ] OpenAI SDK 已就绪 - 支持多厂商 AI 调用
)

echo.

:: ==================== 启动服务 ====================
echo [3/3] 正在启动服务...
echo.
echo ================================================
echo   请在浏览器中打开: http://127.0.0.1:5000
echo   默认账号: admin / admin123
echo   按 Ctrl+C 可停止服务
echo ================================================
echo.

py app.py

pause
