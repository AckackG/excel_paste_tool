@echo off
chcp 65001 > nul

set VENV_DIR_NAME=.venv
set ENTRY=main.py

if exist %VENV_DIR_NAME%\Scripts\activate.bat (
    .\%VENV_DIR_NAME%\Scripts\python.exe %ENTRY%
) else (
    echo 错误：当前目录下找不到名为 "%VENV_DIR_NAME%" 的虚拟环境。
    pause
)
