@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo KindleSnapOCR ビルド中...

if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else (
    echo venvがありません。先にstart.batを実行してください。
    pause
    exit /b 1
)

pip install --quiet pyinstaller

pyinstaller --noconfirm --onedir --windowed ^
    --name "KindleSnapOCR" ^
    --add-data "src;src" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL.Image" ^
    --hidden-import "PIL.ImageGrab" ^
    --hidden-import "PIL.ImageTk" ^
    --hidden-import "winocr" ^
    --hidden-import "fitz" ^
    --hidden-import "pyautogui" ^
    --hidden-import "pygetwindow" ^
    main.py

echo.
echo ビルド完了: dist\KindleSnapOCR\
pause
