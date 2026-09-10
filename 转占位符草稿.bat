@echo off
chcp 936 >nul
title Direct Citation Assistant - 转占位符草稿
set "P=%~1"
if "%P%"=="" (
    echo 用法：把要交给 AI 修改的 Word 文档（.docx）拖到本文件上。
    echo 也可以双击后手动输入完整路径。
    echo.
    set /p "P=请输入文档路径: "
)
if "%P%"=="" (
    echo 未输入路径，退出。
    pause
    exit /b 1
)
echo 正在把引用超链接转为 [CITE:key] 占位符并移除文末参考文献表...
python "%~dp0scripts\insert_refs.py" --docx "%P%" --to-placeholders
echo.
echo 完成！现在可以把这份“纯文本草稿”交给 AI/他人自由修改正文。
echo 对方改完后，双击「恢复引用.bat」即可恢复超链接、编号和文末表。
pause
