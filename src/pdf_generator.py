"""
PDF生成モジュール
キャプチャした画像をPDFに変換する
"""
import os
from pathlib import Path
from typing import List, Optional, Callable
import fitz  # PyMuPDF
from PIL import Image


class PDFGenerator:
    """画像からPDFを生成する"""

    def __init__(self):
        pass

    def images_to_pdf(
        self,
        image_paths: List[str],
        output_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        """
        複数の画像を1つのPDFに変換

        Args:
            image_paths: 画像ファイルパスのリスト
            output_path: 出力PDFファイルパス
            progress_callback: 進捗コールバック (current, total)

        Returns:
            生成されたPDFのパス
        """
        if not image_paths:
            raise ValueError("画像ファイルが指定されていません")

        # 出力ディレクトリを作成
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        # 新しいPDFドキュメントを作成
        doc = fitz.open()

        total = len(image_paths)

        for idx, img_path in enumerate(image_paths):
            # 画像を読み込み
            img = Image.open(img_path)

            # RGBAをRGBに変換（必要な場合）
            if img.mode == 'RGBA':
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            # 一時的にJPEGとして保存（PyMuPDFで読み込むため）
            temp_path = img_path + '.temp.jpg'
            img.save(temp_path, 'JPEG', quality=95)

            # PyMuPDFで画像を読み込み
            img_doc = fitz.open(temp_path)
            rect = img_doc[0].rect

            # PDFページを画像サイズで作成
            page = doc.new_page(width=rect.width, height=rect.height)

            # 画像を挿入
            page.insert_image(rect, filename=temp_path)

            # 一時ファイルを削除
            img_doc.close()
            os.remove(temp_path)

            # 進捗通知
            if progress_callback:
                progress_callback(idx + 1, total)

        # PDFを保存
        doc.save(output_path)
        doc.close()

        return output_path

    def images_to_pdf_direct(
        self,
        image_paths: List[str],
        output_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        """
        複数の画像を1つのPDFに変換（直接変換版）

        Args:
            image_paths: 画像ファイルパスのリスト
            output_path: 出力PDFファイルパス
            progress_callback: 進捗コールバック (current, total)

        Returns:
            生成されたPDFのパス
        """
        if not image_paths:
            raise ValueError("画像ファイルが指定されていません")

        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        doc = fitz.open()
        total = len(image_paths)

        for idx, img_path in enumerate(image_paths):
            # 画像のサイズを取得
            with Image.open(img_path) as img:
                width, height = img.size

            # ページを作成
            page = doc.new_page(width=width, height=height)

            # 画像を直接挿入
            rect = fitz.Rect(0, 0, width, height)
            page.insert_image(rect, filename=img_path)

            if progress_callback:
                progress_callback(idx + 1, total)

        doc.save(output_path)
        doc.close()

        return output_path


class PDFWithOCR:
    """OCR結果を埋め込んだPDFを生成"""

    def __init__(self, tesseract_path: Optional[str] = None):
        self.tesseract_path = tesseract_path

    def create_searchable_pdf(
        self,
        image_paths: List[str],
        ocr_results: List[str],
        output_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> str:
        """
        OCRテキストを透明レイヤーとして埋め込んだPDFを生成

        Args:
            image_paths: 画像ファイルパスのリスト
            ocr_results: 各ページのOCR結果テキストのリスト
            output_path: 出力PDFファイルパス
            progress_callback: 進捗コールバック

        Returns:
            生成されたPDFのパス
        """
        if not image_paths:
            raise ValueError("画像ファイルが指定されていません")

        if len(image_paths) != len(ocr_results):
            raise ValueError("画像数とOCR結果数が一致しません")

        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        doc = fitz.open()
        total = len(image_paths)

        for idx, (img_path, ocr_text) in enumerate(zip(image_paths, ocr_results)):
            with Image.open(img_path) as img:
                width, height = img.size

            # ページを作成
            page = doc.new_page(width=width, height=height)

            # 画像を挿入
            rect = fitz.Rect(0, 0, width, height)
            page.insert_image(rect, filename=img_path)

            # OCRテキストを透明レイヤーとして追加
            if ocr_text.strip():
                # テキストを透明（見えない）色で追加
                # これにより検索可能なPDFになる
                text_point = fitz.Point(10, height - 10)
                page.insert_text(
                    text_point,
                    ocr_text,
                    fontsize=1,  # 非常に小さいフォント
                    color=(1, 1, 1),  # 白（背景と同化）
                    render_mode=3  # 不可視
                )

            if progress_callback:
                progress_callback(idx + 1, total)

        doc.save(output_path)
        doc.close()

        return output_path


if __name__ == '__main__':
    # テスト
    generator = PDFGenerator()
    print("PDF生成モジュール読み込み完了")
