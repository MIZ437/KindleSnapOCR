"""
OCR処理モジュール
Tesseract OCR と manga-ocr に対応
"""
import os
import sys
from pathlib import Path
from typing import List, Optional, Callable
from enum import Enum
import numpy as np
from PIL import Image, ImageFilter, ImageEnhance


class OCREngine(Enum):
    """OCRエンジンの種類"""
    TESSERACT = "tesseract"
    MANGA_OCR = "manga_ocr"


class TextDirection(Enum):
    """テキストの方向"""
    HORIZONTAL = "horizontal"  # 横書き
    VERTICAL = "vertical"      # 縦書き
    MIXED = "mixed"            # 混在


class PreprocessingLevel(Enum):
    """前処理レベル"""
    NONE = "none"              # 前処理なし
    SIMPLE = "simple"          # シンプル（PIL のみ）
    ADVANCED = "advanced"      # 高度（OpenCV使用）


def preprocess_image_simple(image: Image.Image) -> Image.Image:
    """
    シンプルな画像前処理（PIL/numpy のみ）

    Args:
        image: PIL Image

    Returns:
        前処理済みのPIL Image
    """
    # 1. グレースケール変換
    if image.mode != 'L':
        gray = image.convert('L')
    else:
        gray = image

    # 2. コントラスト強調
    enhancer = ImageEnhance.Contrast(gray)
    contrast = enhancer.enhance(1.5)

    # 3. シャープネス強調
    sharpener = ImageEnhance.Sharpness(contrast)
    sharp = sharpener.enhance(2.0)

    # 4. 大津の二値化
    img_array = np.array(sharp)
    threshold = _otsu_threshold(img_array)
    binary = ((img_array > threshold) * 255).astype(np.uint8)

    return Image.fromarray(binary)


def preprocess_image_advanced(image: Image.Image) -> Image.Image:
    """
    高度な画像前処理（OpenCV使用）

    Args:
        image: PIL Image

    Returns:
        前処理済みのPIL Image
    """
    try:
        import cv2
    except ImportError:
        # OpenCVがない場合はシンプル版にフォールバック
        return preprocess_image_simple(image)

    # PIL -> OpenCV
    img_array = np.array(image)
    if len(img_array.shape) == 3:
        if img_array.shape[2] == 4:  # RGBA
            img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
        else:  # RGB
            img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    else:
        gray = img_array

    # 1. CLAHE（適応ヒストグラム均等化）でコントラスト強調
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)

    # 2. 大津の二値化
    _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # 3. ノイズ除去（Non-local Means Denoising）
    denoised = cv2.fastNlMeansDenoising(binary, h=5, templateWindowSize=7, searchWindowSize=21)

    return Image.fromarray(denoised)


def _otsu_threshold(img_array: np.ndarray) -> int:
    """大津の方法で最適な閾値を計算"""
    hist, _ = np.histogram(img_array.flatten(), bins=256, range=(0, 256))
    total = img_array.size

    sum_total = np.sum(np.arange(256) * hist)
    sum_bg = 0
    weight_bg = 0

    max_variance = 0
    threshold = 0

    for i in range(256):
        weight_bg += hist[i]
        if weight_bg == 0:
            continue

        weight_fg = total - weight_bg
        if weight_fg == 0:
            break

        sum_bg += i * hist[i]
        mean_bg = sum_bg / weight_bg
        mean_fg = (sum_total - sum_bg) / weight_fg

        variance = weight_bg * weight_fg * (mean_bg - mean_fg) ** 2

        if variance > max_variance:
            max_variance = variance
            threshold = i

    return threshold


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


def check_manga_ocr_available() -> bool:
    """manga-ocrが利用可能かチェック"""
    try:
        import manga_ocr
        return True
    except ImportError:
        return False


class OCRProcessor:
    """OCR処理クラス（Tesseract / manga-ocr対応）"""

    # Tesseract言語設定
    TESSERACT_LANGUAGES = {
        'jpn': '日本語（横書き）',
        'jpn_vert': '日本語（縦書き）',
        'jpn+jpn_vert': '日本語（横+縦）',
        'eng': '英語',
        'jpn+eng': '日本語+英語',
        'jpn+jpn_vert+eng': '日本語+英語（全対応）',
    }

    # PSMモード
    PSM_MODES = {
        3: '完全自動（デフォルト）',
        4: '単一列（可変サイズ）',
        5: '縦書き単一ブロック',
        6: '単一ブロック（推奨）',
        11: '疎なテキスト',
        13: '単一行',
    }

    def __init__(
        self,
        language: str = 'jpn',
        engine: OCREngine = OCREngine.TESSERACT,
        text_direction: TextDirection = TextDirection.HORIZONTAL,
        preprocessing: PreprocessingLevel = PreprocessingLevel.ADVANCED,
        psm_mode: int = 6
    ):
        """
        Args:
            language: Tesseract言語設定
            engine: OCRエンジン
            text_direction: テキスト方向
            preprocessing: 前処理レベル
            psm_mode: TesseractのPSMモード
        """
        self.language = language
        self.engine = engine
        self.text_direction = text_direction
        self.preprocessing = preprocessing
        self.psm_mode = psm_mode

        self.tesseract_path = find_tesseract()
        self._pytesseract = None
        self._manga_ocr = None

        # Tesseractの初期化
        if self.tesseract_path:
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
                self._pytesseract = pytesseract
            except ImportError:
                pass

        # manga-ocrの初期化（遅延ロード）
        if engine == OCREngine.MANGA_OCR:
            self._init_manga_ocr()

    def _init_manga_ocr(self):
        """manga-ocrの初期化"""
        try:
            from manga_ocr import MangaOcr
            self._manga_ocr = MangaOcr()
        except ImportError:
            pass
        except Exception:
            pass

    def _get_tesseract_config(self) -> str:
        """Tesseract設定文字列を生成"""
        # テキスト方向に応じた言語設定
        if self.text_direction == TextDirection.VERTICAL:
            lang = 'jpn_vert'
            psm = 5  # 縦書き用
        elif self.text_direction == TextDirection.MIXED:
            lang = 'jpn+jpn_vert'
            psm = 3  # 完全自動
        else:
            lang = self.language
            psm = self.psm_mode

        # OEM 3 = LSTM（ニューラルネット、最高精度）
        config = f'--oem 3 --psm {psm} -l {lang}'
        return config

    def is_available(self) -> bool:
        """OCRが利用可能かチェック"""
        if self.engine == OCREngine.MANGA_OCR:
            return self._manga_ocr is not None
        else:
            return self._pytesseract is not None and self.tesseract_path is not None

    def get_engine_name(self) -> str:
        """現在のエンジン名を取得"""
        if self.engine == OCREngine.MANGA_OCR:
            return "manga-ocr"
        else:
            return "Tesseract"

    def _preprocess(self, image: Image.Image) -> Image.Image:
        """前処理を適用"""
        if self.preprocessing == PreprocessingLevel.NONE:
            return image
        elif self.preprocessing == PreprocessingLevel.ADVANCED:
            return preprocess_image_advanced(image)
        else:
            return preprocess_image_simple(image)

    def process_image(self, image_path: str, use_preprocessing: bool = True) -> str:
        """画像からテキストを抽出"""
        if not self.is_available():
            return ""

        image = Image.open(image_path)

        if self.engine == OCREngine.MANGA_OCR:
            # manga-ocrは前処理不要（モデルが対応）
            text = self._manga_ocr(image)
        else:
            # Tesseractは前処理が有効
            if use_preprocessing:
                image = self._preprocess(image)

            config = self._get_tesseract_config()
            text = self._pytesseract.image_to_string(image, config=config)

        return text

    def process_pil_image(self, image: Image.Image, use_preprocessing: bool = True) -> str:
        """PIL Imageからテキストを抽出"""
        if not self.is_available():
            return ""

        if self.engine == OCREngine.MANGA_OCR:
            text = self._manga_ocr(image)
        else:
            if use_preprocessing:
                image = self._preprocess(image)

            config = self._get_tesseract_config()
            text = self._pytesseract.image_to_string(image, config=config)

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
                engine_name = self.get_engine_name()
                progress_callback(idx + 1, total, f"{engine_name}: {Path(img_path).name}")

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

    def process_pdf(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        dpi: int = 200
    ) -> List[str]:
        """
        PDFファイルからテキストを抽出

        Args:
            pdf_path: PDFファイルパス
            progress_callback: (current, total, status) コールバック
            dpi: 画像変換時の解像度

        Returns:
            各ページのOCRテキストのリスト
        """
        if not self.is_available():
            return []

        import fitz  # PyMuPDF
        import io

        results = []
        doc = fitz.open(pdf_path)
        total = len(doc)

        try:
            for page_num in range(total):
                if progress_callback:
                    engine_name = self.get_engine_name()
                    progress_callback(page_num + 1, total, f"{engine_name}: {page_num + 1}/{total}ページ")

                page = doc[page_num]
                # ページを画像に変換
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                pix = page.get_pixmap(matrix=mat)

                # PIL Imageに変換
                img_data = pix.tobytes("png")
                image = Image.open(io.BytesIO(img_data))

                # OCR実行
                try:
                    text = self.process_pil_image(image)
                    results.append(text)
                except Exception as e:
                    results.append(f"[OCR Error on page {page_num + 1}: {str(e)}]")

        finally:
            doc.close()

        return results

    def process_pdf_to_file(
        self,
        pdf_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        dpi: int = 200
    ) -> str:
        """
        PDFファイルをOCR処理してテキストファイルに保存

        Args:
            pdf_path: PDFファイルパス
            output_path: 出力テキストファイルパス（省略時はPDFと同じ場所に_ocr.txt）
            progress_callback: (current, total, status) コールバック
            dpi: 画像変換時の解像度

        Returns:
            出力ファイルパス
        """
        if output_path is None:
            base = os.path.splitext(pdf_path)[0]
            output_path = f"{base}_ocr.txt"

        results = self.process_pdf(pdf_path, progress_callback, dpi)
        self.save_ocr_results(results, output_path)

        return output_path


class MangaOCRProcessor:
    """manga-ocr専用プロセッサ（シングルトン）"""
    _instance = None
    _mocr = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def get_ocr(self):
        """manga-ocrインスタンスを取得（遅延初期化）"""
        if self._mocr is None:
            try:
                from manga_ocr import MangaOcr
                self._mocr = MangaOcr()
            except ImportError:
                raise ImportError(
                    "manga-ocrがインストールされていません。\n"
                    "pip install manga-ocr を実行してください。"
                )
        return self._mocr

    def is_available(self) -> bool:
        """manga-ocrが利用可能かチェック"""
        return check_manga_ocr_available()

    def release(self):
        """メモリ解放"""
        self._mocr = None
        import gc
        gc.collect()

    def process(self, image: Image.Image) -> str:
        """画像からテキストを抽出"""
        mocr = self.get_ocr()
        return mocr(image)
