@echo off
chcp 65001 > nul

REM 仮想環境があれば有効化
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM アプリケーションを実行
python main.py
