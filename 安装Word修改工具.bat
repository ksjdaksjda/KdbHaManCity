@echo off
chcp 65001 >nul
title Word论文修改工具 - 安装
echo ============================================
echo   Word论文修改工具 - 一键安装
echo ============================================
echo.
echo 正在检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python
    echo 下载: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo Python 已就绪
echo.
echo 正在安装依赖包 (python-docx + openai)...
pip install python-docx openai -q
echo 依赖已安装
echo.
echo ============================================
echo   安装完成！使用方式：
echo.
echo   双击 word_reviser.py 运行
echo   或者命令行: python word_reviser.py
echo ============================================
echo.
pause
