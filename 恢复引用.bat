@echo off
chcp 936 >nul
title Direct Citation Assistant - 恢复引用
set "P=%~1"
if "%P%"=="" (
    echo 用法：把 AI 改完的 Word 文档（.docx）拖到本文件上。
    echo 也可以双击后手动输入完整路径。
    echo.
    set /p "P=请输入文档路径: "
)
if "%P%"=="" (
    echo 未输入路径，退出。
    pause
    exit /b 1
)
set "REFS=%~dp1..\引用目录\refs.csv"
if not exist "%REFS%" set "REFS=%~dp1refs.csv"
if not exist "%REFS%" (
    echo 未找到引用表 refs.csv（默认尝试文档同目录或上级“引用目录”）。
    set /p "REFS=请手动输入 refs.csv 完整路径: "
)
echo 正在恢复引用：占位符转为超链接、编号重排、文末表重建...
python "%~dp0scripts\insert_refs.py" --docx "%P%" --refs "%REFS%" --mapping
echo.
echo 完成！引用已恢复。
echo 上方若有“正文引用 - 文献对照表”，请逐处核对引用是否真实支撑该句。
pause
