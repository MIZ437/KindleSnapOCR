"""
OCR処理モジュール
Tesseract OCRがインストールされていれば使用、なければスキップ
"""
import os
import sys
from pathlib import Path
from typing import List, Optional, Callable


def find_tesseract() -> Optional[str]:
    """システムからTesseractを探す"""
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        os.path.expanduser(r'~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
    ]

    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    possible_paths.insert(0, os.path.join(exe_dir, 'tesseract', 'tesseract.exe'))
    possible_paths.insert(0, os.path.join(exe_dir, '..', 'tesseract', 'tesseract.exe'))

    for path in possible_paths:
        if os.path.exists(path):
            return path

    return None


class OCRProcessor:
    """Tesseract OCRを使用したテキスト抽出（オプション機能）"""

    LANGUAGES = {
        'jpn': '日本語',
        'eng': '英語',
        'jpn+eng': '日本語+英語',
    }

    def __init__(self, language: str = 'jpn'):
        self.language = language
        self.tesseract_path = find_tesseract()
        self._pytesseract = None

        if self.tesseract_path:
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
                self._pytesseract = pytesseract
            except ImportError:
                pass

    def is_available(self) -> bool:
        """OCRが利用可能かチェック"""
        return self._pytesseract is not None and self.tesseract_path is not None

    def process_image(self, image_path: str) -> str:
        """画像からテキストを抽出"""
        if not self.is_available():
            return ""

        from PIL import Image
        image = Image.open(image_path)
        text = self._pytesseract.image_to_string(image, lang=self.language)
        return text

    def process_images(
        self,
        image_paths: List[str],
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[str]:
        """複数の画像からテキストを抽出"""
        if not self.is_available():
            return [""] * len(image_paths)

        results = []
        total = len(image_paths)

        for idx, img_path in enumerate(image_paths):
            if progress_callback:
                progress_callback(idx + 1, total, f"OCR: {Path(img_path).name}")

            try:
                text = self.process_image(img_path)
                results.append(text)
            except Exception as e:
                results.append(f"[OCR Error: {str(e)}]")

        return results

    def save_ocr_results(
        self,
        ocr_results: List[str],
        output_path: str,
        page_separator: str = "\n\n--- Page {page} ---\n\n"
    ):
        """OCR結果をテキストファイルに保存"""
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, 'w', encoding='utf-8') as f:
            for idx, text in enumerate(ocr_results):
                if idx > 0:
                    f.write(page_separator.format(page=idx + 1))
                f.write(text)
