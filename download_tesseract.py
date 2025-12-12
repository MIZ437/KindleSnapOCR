"""
Tesseract OCR ポータブル版ダウンローダー
"""
import os
import sys
import zipfile
import urllib.request
import shutil
from pathlib import Path


def download_file(url, dest_path, desc="ダウンロード中"):
    """ファイルをダウンロード（プログレス表示付き）"""
    def progress_hook(block_num, block_size, total_size):
        downloaded = block_num * block_size
        if total_size > 0:
            percent = min(100, downloaded * 100 // total_size)
            bar_len = 30
            filled = int(bar_len * percent // 100)
            bar = '=' * filled + '-' * (bar_len - filled)
            print(f'\r{desc}: [{bar}] {percent}%', end='', flush=True)

    urllib.request.urlretrieve(url, dest_path, progress_hook)
    print()


def main():
    """Tesseract OCRをダウンロードしてセットアップ"""
    script_dir = Path(__file__).parent
    tesseract_dir = script_dir / "tesseract"
    tessdata_dir = tesseract_dir / "tessdata"

    # 既にインストール済みならスキップ
    if (tesseract_dir / "tesseract.exe").exists():
        print("Tesseract OCRは既にインストールされています")
        return

    print("Tesseract OCR ポータブル版をセットアップします...")
    print()

    # 一時ディレクトリ
    temp_dir = script_dir / "temp_tesseract"
    temp_dir.mkdir(exist_ok=True)

    try:
        # Tesseractポータブル版のダウンロード
        # UB-Mannheim のポータブルZIPを使用
        tesseract_url = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe"

        # 代替: tesseract-ocr-w64 をそのまま使う方式（インストーラーを使わない）
        # ポータブル用にビルド済みのものを使用
        zip_url = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe"

        # tessdata (言語データ) を個別にダウンロード
        tessdata_urls = {
            "eng.traineddata": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata",
            "jpn.traineddata": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/jpn.traineddata",
            "jpn_vert.traineddata": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/jpn_vert.traineddata",
        }

        # Tesseract本体（ポータブル版 zip）
        # 注: 公式にはzipが無いため、7zで展開するか別の方法を使う
        # ここではwindows-tesseract-ocrのポータブルビルドを使用

        print("代替方法: システムにTesseractをインストールします...")
        print()

        # インストーラーをダウンロードして実行（ユーザーにインストールしてもらう）
        installer_path = temp_dir / "tesseract_installer.exe"

        print("Tesseractインストーラーをダウンロード中...")
        download_file(
            "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe",
            str(installer_path),
            "インストーラー"
        )

        print()
        print("=" * 50)
        print("Tesseractインストーラーを起動します。")
        print()
        print("【重要】インストール手順:")
        print("1. 「Next」で進む")
        print("2. 「I Agree」で同意")
        print("3. 「Additional language data」を展開")
        print("4. 「Japanese」と「Japanese (vertical)」にチェック")
        print("5. 「Next」→「Install」")
        print("=" * 50)
        print()

        # インストーラーを起動
        os.startfile(str(installer_path))

        input("インストールが完了したらEnterキーを押してください...")

        # インストール確認
        if os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
            print("Tesseractのインストールを確認しました。")
        else:
            print("警告: Tesseractが見つかりません。OCR機能が使えない可能性があります。")

    except Exception as e:
        print(f"エラー: {e}")
        print()
        print("手動でTesseractをインストールしてください:")
        print("https://github.com/UB-Mannheim/tesseract/wiki")

    finally:
        # 一時ファイル削除
        if temp_dir.exists():
            shutil.rmtree(temp_dir, ignore_errors=True)


def setup_portable():
    """ポータブル版のセットアップ（tesseractフォルダに展開）"""
    script_dir = Path(__file__).parent
    tesseract_dir = script_dir / "tesseract"
    tessdata_dir = tesseract_dir / "tessdata"

    tesseract_dir.mkdir(exist_ok=True)
    tessdata_dir.mkdir(exist_ok=True)

    # 言語データのダウンロード
    tessdata_urls = {
        "eng.traineddata": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata",
        "jpn.traineddata": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/jpn.traineddata",
        "jpn_vert.traineddata": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/jpn_vert.traineddata",
    }

    for filename, url in tessdata_urls.items():
        dest = tessdata_dir / filename
        if not dest.exists():
            print(f"ダウンロード中: {filename}")
            download_file(url, str(dest), filename)


if __name__ == "__main__":
    main()
