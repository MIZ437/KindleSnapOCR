"""Tesseract OCRインストールスクリプト（start.ps1から呼び出される）"""
import sys
sys.path.insert(0, 'src')
from tesseract_installer import download_and_install_tesseract, ensure_japanese_installed

def progress(status, current, total):
    print('  ' + status)

def lang_progress(status):
    print('  ' + status)

# Tesseract本体のインストール
success = download_and_install_tesseract(progress)

if success:
    # 日本語言語パックのインストール
    print('')
    print('  Installing Japanese language pack...')
    ensure_japanese_installed(lang_progress)

sys.exit(0 if success else 1)
