@echo off
chcp 65001 >nul
title Word论文修改工具
cd /d R:\ThesisWriter

echo ============================================
echo   Word论文修改工具
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python
    echo 下载: https://www.python.org/downloads/
    pause
    exit /b 1
)

python -c "import docx, openai" 2>nul
if errorlevel 1 (
    echo 正在安装依赖...
    pip install python-docx openai -q
)

python word_reviser.py
pause
