@echo off
chcp 65001 > nul
echo ==========================================
echo KindleSnapOCR セットアップ
echo ==========================================
echo.

REM Pythonのバージョンチェック
python --version >nul 2>&1
if errorlevel 1 (
    echo エラー: Pythonがインストールされていません
    echo https://www.python.org/downloads/ からインストールしてください
    pause
    exit /b 1
)

echo Pythonが検出されました:
python --version
echo.

REM 仮想環境の作成
echo 仮想環境を作成しています...
python -m venv venv
if errorlevel 1 (
    echo エラー: 仮想環境の作成に失敗しました
    pause
    exit /b 1
)

REM 仮想環境の有効化
echo 仮想環境を有効化しています...
call venv\Scripts\activate.bat

REM pipのアップグレード
echo pipをアップグレードしています...
python -m pip install --upgrade pip

REM 依存関係のインストール
echo 依存関係をインストールしています...
pip install -r requirements.txt

echo.
echo ==========================================
echo セットアップが完了しました
echo ==========================================
echo.
echo 次のステップ:
echo 1. Tesseract-OCRをインストール（OCR機能を使う場合）
echo    ダウンロード: https://github.com/UB-Mannheim/tesseract/wiki
echo    インストール時に「Additional language data」で「Japanese」を選択
echo.
echo 2. アプリを実行:
echo    run.bat をダブルクリック
echo.
echo 3. EXEファイルを作成（ポータブル版）:
echo    build.bat をダブルクリック
echo.
pause
