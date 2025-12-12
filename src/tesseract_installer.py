"""
Tesseract OCR 自動インストーラー
UB Mannheim版Tesseractをダウンロードしてインストール
"""
import os
import sys
import urllib.request
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Callable

# Tesseract OCR ダウンロードURL (UB Mannheim)
TESSERACT_URL = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe"
TESSERACT_FILENAME = "tesseract-ocr-setup.exe"

# 言語パックURL (tessdata)
LANG_DATA_URLS = {
    'jpn': 'https://github.com/tesseract-ocr/tessdata/raw/main/jpn.traineddata',
    'jpn_vert': 'https://github.com/tesseract-ocr/tessdata/raw/main/jpn_vert.traineddata',
}


def is_tesseract_installed() -> bool:
    """Tesseractがインストールされているかチェック"""
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        os.path.expanduser(r'~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return True
    return False


def get_tesseract_path() -> Optional[str]:
    """Tesseractのパスを取得"""
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        os.path.expanduser(r'~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None


def get_tessdata_path() -> Optional[str]:
    """tessdataフォルダのパスを取得"""
    tesseract_path = get_tesseract_path()
    if tesseract_path:
        tessdata = os.path.join(os.path.dirname(tesseract_path), 'tessdata')
        if os.path.exists(tessdata):
            return tessdata
    return None


def is_language_installed(lang: str) -> bool:
    """指定言語がインストールされているかチェック"""
    tessdata = get_tessdata_path()
    if tessdata:
        return os.path.exists(os.path.join(tessdata, f'{lang}.traineddata'))
    return False


def download_language(
    lang: str,
    progress_callback: Optional[Callable[[str], None]] = None
) -> bool:
    """
    言語パックをダウンロード

    Args:
        lang: 言語コード (jpn, jpn_vert など)
        progress_callback: ステータスを受け取るコールバック

    Returns:
        成功したらTrue
    """
    if lang not in LANG_DATA_URLS:
        return False

    tessdata = get_tessdata_path()
    if not tessdata:
        return False

    url = LANG_DATA_URLS[lang]
    dest_path = os.path.join(tessdata, f'{lang}.traineddata')

    if os.path.exists(dest_path):
        return True

    try:
        if progress_callback:
            progress_callback(f'Downloading {lang}...')
        urllib.request.urlretrieve(url, dest_path)
        return True
    except Exception as e:
        if progress_callback:
            progress_callback(f'Error: {str(e)}')
        return False


def ensure_japanese_installed(
    progress_callback: Optional[Callable[[str], None]] = None
) -> bool:
    """日本語言語パックがインストールされていなければダウンロード"""
    if not is_tesseract_installed():
        return False

    success = True
    for lang in ['jpn', 'jpn_vert']:
        if not is_language_installed(lang):
            if progress_callback:
                progress_callback(f'Installing {lang} language pack...')
            if not download_language(lang, progress_callback):
                success = False
        else:
            if progress_callback:
                progress_callback(f'{lang} already installed')

    return success


def download_tesseract(
    progress_callback: Optional[Callable[[int, int], None]] = None
) -> str:
    """
    Tesseractインストーラーをダウンロード

    Args:
        progress_callback: (downloaded_bytes, total_bytes) を受け取るコールバック

    Returns:
        ダウンロードしたファイルのパス
    """
    temp_dir = tempfile.gettempdir()
    installer_path = os.path.join(temp_dir, TESSERACT_FILENAME)

    # 既にダウンロード済みなら再利用
    if os.path.exists(installer_path):
        file_size = os.path.getsize(installer_path)
        if file_size > 50_000_000:  # 50MB以上ならOK
            return installer_path

    def reporthook(block_num, block_size, total_size):
        if progress_callback and total_size > 0:
            downloaded = block_num * block_size
            progress_callback(downloaded, total_size)

    urllib.request.urlretrieve(TESSERACT_URL, installer_path, reporthook)
    return installer_path


def install_tesseract(
    installer_path: str,
    include_japanese: bool = True
) -> bool:
    """
    Tesseractをサイレントインストール

    Args:
        installer_path: インストーラーのパス
        include_japanese: 日本語言語パックを含めるか

    Returns:
        インストール成功したらTrue
    """
    if not os.path.exists(installer_path):
        return False

    # サイレントインストールコマンド
    # /S: サイレントモード
    # /D: インストール先（省略でデフォルト）
    cmd = [installer_path, '/S']

    try:
        # 管理者権限で実行
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=300  # 5分タイムアウト
        )

        # インストール完了を確認
        import time
        for _ in range(30):  # 最大30秒待機
            if is_tesseract_installed():
                return True
            time.sleep(1)

        return is_tesseract_installed()

    except subprocess.TimeoutExpired:
        return False
    except Exception:
        return False


def download_and_install_tesseract(
    progress_callback: Optional[Callable[[str, int, int], None]] = None
) -> bool:
    """
    Tesseractをダウンロードしてインストール

    Args:
        progress_callback: (status, current, total) を受け取るコールバック

    Returns:
        成功したらTrue
    """
    if is_tesseract_installed():
        if progress_callback:
            progress_callback("既にインストール済み", 100, 100)
        return True

    try:
        # ダウンロード
        def download_progress(downloaded, total):
            if progress_callback:
                percent = int(downloaded * 100 / total) if total > 0 else 0
                mb_downloaded = downloaded / (1024 * 1024)
                mb_total = total / (1024 * 1024)
                progress_callback(
                    f"ダウンロード中: {mb_downloaded:.1f}MB / {mb_total:.1f}MB",
                    downloaded,
                    total
                )

        if progress_callback:
            progress_callback("ダウンロードを開始...", 0, 100)

        installer_path = download_tesseract(download_progress)

        # インストール
        if progress_callback:
            progress_callback("インストール中... (管理者権限が必要な場合があります)", 0, 0)

        success = install_tesseract(installer_path)

        if success:
            if progress_callback:
                progress_callback("インストール完了", 100, 100)
            # インストーラーを削除
            try:
                os.remove(installer_path)
            except:
                pass
        else:
            if progress_callback:
                progress_callback("インストール失敗", 0, 100)

        return success

    except Exception as e:
        if progress_callback:
            progress_callback(f"エラー: {str(e)}", 0, 100)
        return False


if __name__ == '__main__':
    # テスト
    def progress(status, current, total):
        if total > 0:
            percent = int(current * 100 / total)
            print(f"\r{status} [{percent}%]", end='', flush=True)
        else:
            print(f"\r{status}", end='', flush=True)

    print("Tesseract OCR インストーラー")
    print("=" * 40)

    if is_tesseract_installed():
        print(f"Tesseractは既にインストールされています: {get_tesseract_path()}")
    else:
        print("Tesseractをインストールします...")
        success = download_and_install_tesseract(progress)
        print()
        if success:
            print("インストール完了!")
        else:
            print("インストールに失敗しました")
