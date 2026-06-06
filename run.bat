@echo off
chcp 65001 >nul
title 论文工具集
cd /d "%~dp0"

:menu
cls
echo ============================================
echo   论文工具集
echo ============================================
echo.
echo   1. 模板分析器 - 读取学校Word模板，提取排版格式
echo   2. 论文重写器 - AI重写成�~^�~F文，只改文字
echo   3. 模板填充器 - 将AI写作内容灌入模板（排版不动）
echo   4. 退出
echo.
set /p choice="请选择 (1/2/3/4): "

if "%choice%"=="1" goto analyzer
if "%choice%"=="2" goto rewriter
if "%choice%"=="3" goto filler
if "%choice%"=="4" exit
goto menu

:analyzer
echo [*] 启动模板分析器...
python --version >nul 2>&1 || (echo [X] 未找到Python && pause && goto menu)
python -c "import docx" 2>nul || pip install python-docx -q
python template_analyzer.py
goto menu

:rewriter
echo [*] 启动论文重写器...
python --version >nul 2>&1 || (echo [X] 未找到Python && pause && goto menu)
python -c "import docx,openai" 2>nul || pip install python-docx openai -q
python word_reviser.py
goto menu

:filler
echo [*] 启动模板填充器...
python --version >nul 2>&1 || (echo [X] 未找到Python && pause && goto menu)
python -c "import docx" 2>nul || pip install python-docx -q
python fill_template.py
goto menu
