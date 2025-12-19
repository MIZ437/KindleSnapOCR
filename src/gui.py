"""
GUIモジュール
メインウィンドウとユーザーインターフェース
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from pathlib import Path
from datetime import datetime

# DPI awareness
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    pass


class MainWindow:
    """メインウィンドウ"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("KindleSnapOCR - Kindle本PDF化ツール")
        self.root.geometry("650x800")
        self.root.resizable(True, True)

        # 状態変数
        self.capture_region = None
        self.is_capturing = False
        self.stop_flag = False
        self.captured_files = []

        # 設定
        self.page_direction = tk.StringVar(value='right')
        self.stop_mode = tk.StringVar(value='manual')  # 'auto', 'manual', 'pages'
        self.total_pages = tk.StringVar(value='')
        self.auto_detect_count = tk.StringVar(value='10')  # 自動検出回数
        self.delay_time = tk.StringVar(value='0.5')
        self.ocr_language = tk.StringVar(value='jpn')
        self.output_folder = tk.StringVar(value='')
        self.book_title = tk.StringVar(value='')
        self.enable_ocr = tk.BooleanVar(value=True)
        self.privacy_mode = tk.BooleanVar(value=False)
        self.privacy_controller = None

        # OCR詳細設定
        self.ocr_engine = tk.StringVar(value='tesseract')
        self.text_direction = tk.StringVar(value='horizontal')
        self.preprocessing_level = tk.StringVar(value='advanced')

        self._setup_ui()
        self._set_default_output()
        self._check_ocr()

    def _setup_ui(self):
        """UIをセットアップ"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 書籍情報 ===
        book_frame = ttk.LabelFrame(main_frame, text="書籍情報", padding="10")
        book_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(book_frame, text="書籍タイトル:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(book_frame, textvariable=self.book_title, width=50).grid(row=0, column=1, sticky=tk.EW, padx=(5, 0))
        book_frame.columnconfigure(1, weight=1)

        # === キャプチャ設定 ===
        capture_frame = ttk.LabelFrame(main_frame, text="キャプチャ設定", padding="10")
        capture_frame.pack(fill=tk.X, pady=(0, 10))

        region_frame = ttk.Frame(capture_frame)
        region_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(region_frame, text="キャプチャ範囲を選択", command=self._select_region).pack(side=tk.LEFT)
        self.region_label = ttk.Label(region_frame, text="未選択", foreground="red")
        self.region_label.pack(side=tk.LEFT, padx=(10, 0))

        direction_frame = ttk.Frame(capture_frame)
        direction_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(direction_frame, text="ページ送り方向:").pack(side=tk.LEFT)
        ttk.Radiobutton(direction_frame, text="→ 右 (横書き)", variable=self.page_direction, value='right').pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(direction_frame, text="← 左 (縦書き)", variable=self.page_direction, value='left').pack(side=tk.LEFT, padx=(10, 0))

        # 停止モード選択
        stop_frame = ttk.Frame(capture_frame)
        stop_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(stop_frame, text="停止方法:").pack(side=tk.LEFT)
        ttk.Radiobutton(stop_frame, text="手動 (ESC+通知)", variable=self.stop_mode, value='manual', command=self._toggle_page_input).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(stop_frame, text="自動検出", variable=self.stop_mode, value='auto', command=self._toggle_page_input).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(stop_frame, text="ページ数指定", variable=self.stop_mode, value='pages', command=self._toggle_page_input).pack(side=tk.LEFT, padx=(10, 0))

        # 自動検出回数とページ数の入力
        param_frame = ttk.Frame(capture_frame)
        param_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(param_frame, text="自動検出回数:").pack(side=tk.LEFT)
        self.detect_combo = ttk.Combobox(param_frame, textvariable=self.auto_detect_count,
                                          values=[str(i) for i in range(5, 21)], width=5, state='readonly')
        self.detect_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.detect_desc_label = ttk.Label(param_frame, text=f"(同じ画像が連続{self.auto_detect_count.get()}回で停止)", foreground="gray")
        self.detect_desc_label.pack(side=tk.LEFT, padx=(5, 0))
        self.detect_combo.bind('<<ComboboxSelected>>', self._update_detect_desc)

        ttk.Label(param_frame, text="総ページ数:").pack(side=tk.LEFT, padx=(20, 0))
        self.pages_entry = ttk.Entry(param_frame, textvariable=self.total_pages, width=10)
        self.pages_entry.pack(side=tk.LEFT, padx=(5, 0))
        self.pages_entry.config(state='disabled')

        delay_frame = ttk.Frame(capture_frame)
        delay_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(delay_frame, text="ページ送り後の待機時間(秒):").pack(side=tk.LEFT)
        ttk.Entry(delay_frame, textvariable=self.delay_time, width=10).pack(side=tk.LEFT, padx=(5, 0))

        # プライバシーモード
        privacy_frame = ttk.Frame(capture_frame)
        privacy_frame.pack(fill=tk.X)

        ttk.Checkbutton(privacy_frame, text="プライバシーモード (キャプチャ領域を黒で隠す)", variable=self.privacy_mode).pack(side=tk.LEFT)

        # === OCR設定 ===
        ocr_frame = ttk.LabelFrame(main_frame, text="OCR設定（文字認識）", padding="10")
        ocr_frame.pack(fill=tk.X, pady=(0, 10))

        # 表示名と内部値のマッピング
        self._engine_map = {
            'Tesseract（軽量・汎用）': 'tesseract',
            'manga-ocr（高精度・日本語特化）': 'manga_ocr'
        }
        self._direction_map = {
            '横書き': 'horizontal',
            '縦書き': 'vertical',
            '自動判定': 'mixed'
        }
        self._preproc_map = {
            'なし': 'none',
            '標準': 'simple',
            '高精度（推奨）': 'advanced'
        }

        # OCR有効化とステータス
        ocr_row1 = ttk.Frame(ocr_frame)
        ocr_row1.pack(fill=tk.X)

        self.ocr_check = ttk.Checkbutton(ocr_row1, text="OCR処理を行う（PDFからテキストを抽出）", variable=self.enable_ocr)
        self.ocr_check.pack(side=tk.LEFT)

        self.ocr_status_label = ttk.Label(ocr_row1, text="")
        self.ocr_status_label.pack(side=tk.LEFT, padx=(10, 0))

        self.install_tesseract_btn = ttk.Button(ocr_row1, text="Tesseractをインストール", command=self._install_tesseract)
        self.install_tesseract_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.install_tesseract_btn.pack_forget()  # 初期状態では非表示

        # OCRエンジン選択
        ocr_row2 = ttk.Frame(ocr_frame)
        ocr_row2.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(ocr_row2, text="認識エンジン:").pack(side=tk.LEFT)
        self._engine_display = tk.StringVar(value='Tesseract（軽量・汎用）')
        self.engine_combo = ttk.Combobox(ocr_row2, textvariable=self._engine_display,
                                          values=list(self._engine_map.keys()), width=28, state='readonly')
        self.engine_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.engine_combo.bind('<<ComboboxSelected>>', self._on_engine_change)

        self.engine_status_label = ttk.Label(ocr_row2, text="", foreground="gray")
        self.engine_status_label.pack(side=tk.LEFT, padx=(10, 0))

        self.install_manga_ocr_btn = ttk.Button(ocr_row2, text="インストール", command=self._install_manga_ocr)
        self.install_manga_ocr_btn.pack(side=tk.LEFT, padx=(5, 0))
        self.install_manga_ocr_btn.pack_forget()

        # テキスト方向と前処理
        ocr_row3 = ttk.Frame(ocr_frame)
        ocr_row3.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(ocr_row3, text="本の種類:").pack(side=tk.LEFT)
        self._direction_display = tk.StringVar(value='横書き')
        self.direction_combo = ttk.Combobox(ocr_row3, textvariable=self._direction_display,
                                             values=list(self._direction_map.keys()), width=10, state='readonly')
        self.direction_combo.pack(side=tk.LEFT, padx=(5, 0))

        ttk.Label(ocr_row3, text="精度:").pack(side=tk.LEFT, padx=(15, 0))
        self._preproc_display = tk.StringVar(value='高精度（推奨）')
        self.preproc_combo = ttk.Combobox(ocr_row3, textvariable=self._preproc_display,
                                           values=list(self._preproc_map.keys()), width=14, state='readonly')
        self.preproc_combo.pack(side=tk.LEFT, padx=(5, 0))

        # === テキスト抽出（OCR不要） ===
        extract_frame = ttk.LabelFrame(main_frame, text="テキスト抽出（PDF・Wordから直接抽出、高速・高精度）", padding="10")
        extract_frame.pack(fill=tk.X, pady=(0, 10))

        extract_row1 = ttk.Frame(extract_frame)
        extract_row1.pack(fill=tk.X)

        self.extract_pdf_btn = ttk.Button(extract_row1, text="PDFからテキスト抽出", command=self._extract_pdf_text)
        self.extract_pdf_btn.pack(side=tk.LEFT)

        self.extract_word_btn = ttk.Button(extract_row1, text="Wordからテキスト抽出", command=self._extract_word_text)
        self.extract_word_btn.pack(side=tk.LEFT, padx=(10, 0))

        self.extract_status = ttk.Label(extract_row1, text="")
        self.extract_status.pack(side=tk.LEFT, padx=(10, 0))

        extract_desc = ttk.Label(extract_frame, text="※ テキスト付きPDF・Word文書から直接抽出（OCRより高速・正確）", foreground="gray")
        extract_desc.pack(anchor=tk.W, pady=(5, 0))

        # === 既存PDF/画像のOCR ===
        pdf_ocr_frame = ttk.LabelFrame(main_frame, text="OCR（スキャン画像・写真から文字認識）", padding="10")
        pdf_ocr_frame.pack(fill=tk.X, pady=(0, 10))

        pdf_ocr_row1 = ttk.Frame(pdf_ocr_frame)
        pdf_ocr_row1.pack(fill=tk.X)

        self.pdf_ocr_btn = ttk.Button(pdf_ocr_row1, text="スキャンPDFをOCR", command=self._ocr_existing_pdf)
        self.pdf_ocr_btn.pack(side=tk.LEFT)

        self.image_ocr_btn = ttk.Button(pdf_ocr_row1, text="画像をOCR", command=self._ocr_existing_images)
        self.image_ocr_btn.pack(side=tk.LEFT, padx=(10, 0))

        self.pdf_ocr_status = ttk.Label(pdf_ocr_row1, text="")
        self.pdf_ocr_status.pack(side=tk.LEFT, padx=(10, 0))

        pdf_ocr_desc = ttk.Label(pdf_ocr_frame, text="※ スキャンしたPDFや写真から文字を読み取ります（時間がかかります）", foreground="gray")
        pdf_ocr_desc.pack(anchor=tk.W, pady=(5, 0))

        # === 出力設定 ===
        output_frame = ttk.LabelFrame(main_frame, text="出力設定", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        output_path_frame = ttk.Frame(output_frame)
        output_path_frame.pack(fill=tk.X)

        ttk.Label(output_path_frame, text="出力先:").pack(side=tk.LEFT)
        ttk.Entry(output_path_frame, textvariable=self.output_folder, width=40).pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        ttk.Button(output_path_frame, text="参照", command=self._browse_output).pack(side=tk.LEFT, padx=(5, 0))

        # === 操作ボタン ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_button = ttk.Button(button_frame, text="キャプチャ開始", command=self._start_capture)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_button = ttk.Button(button_frame, text="中断 (ESC)", command=self._stop_capture, state='disabled')
        self.stop_button.pack(side=tk.LEFT)

        # === 進捗 ===
        progress_frame = ttk.LabelFrame(main_frame, text="進捗", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        # 初期状態は非表示（処理開始時に表示）

        self.status_label = ttk.Label(progress_frame, text="待機中...")
        self.status_label.pack(anchor=tk.W)

        log_frame = ttk.Frame(progress_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.log_text = tk.Text(log_frame, height=10, state='disabled')
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.root.bind('<Escape>', lambda e: self._stop_capture())

    def _set_default_output(self):
        default_path = os.path.join(os.path.expanduser('~'), 'Documents', 'KindleSnapOCR')
        self.output_folder.set(default_path)

    def _check_ocr(self):
        """OCRエンジンの状態をチェック"""
        from .ocr_processor import find_tesseract, check_manga_ocr_available

        # Tesseractチェック
        tesseract_path = find_tesseract()
        if tesseract_path:
            self.ocr_status_label.config(text="(Tesseract検出済)", foreground="green")
            self.ocr_check.config(state='normal')
            self.install_tesseract_btn.pack_forget()
        else:
            self.ocr_status_label.config(text="(Tesseract未検出)", foreground="orange")
            self.enable_ocr.set(False)
            self.ocr_check.config(state='disabled')
            self.install_tesseract_btn.pack(side=tk.LEFT, padx=(10, 0))

        # manga-ocrチェック
        self._update_manga_ocr_status()

    def _get_engine_value(self) -> str:
        """表示名から内部値を取得"""
        return self._engine_map.get(self._engine_display.get(), 'tesseract')

    def _get_direction_value(self) -> str:
        """表示名から内部値を取得"""
        return self._direction_map.get(self._direction_display.get(), 'horizontal')

    def _get_preproc_value(self) -> str:
        """表示名から内部値を取得"""
        return self._preproc_map.get(self._preproc_display.get(), 'advanced')

    def _update_manga_ocr_status(self):
        """manga-ocrの状態を更新"""
        from .ocr_processor import check_manga_ocr_available

        if check_manga_ocr_available():
            self.engine_status_label.config(text="利用可能", foreground="green")
            self.install_manga_ocr_btn.pack_forget()
        else:
            engine = self._get_engine_value()
            if engine == 'manga_ocr':
                self.engine_status_label.config(text="未インストール", foreground="orange")
                self.install_manga_ocr_btn.pack(side=tk.LEFT, padx=(5, 0))
            else:
                self.engine_status_label.config(text="")
                self.install_manga_ocr_btn.pack_forget()

    def _on_engine_change(self, event=None):
        """OCRエンジン変更時の処理"""
        engine = self._get_engine_value()

        if engine == 'manga_ocr':
            # manga-ocrは前処理不要、テキスト方向自動
            self.preproc_combo.config(state='disabled')
            self.direction_combo.config(state='disabled')
            self._preproc_display.set('なし')
            self._direction_display.set('自動判定')
        else:
            self.preproc_combo.config(state='readonly')
            self.direction_combo.config(state='readonly')

        self._update_manga_ocr_status()

    def _install_tesseract(self):
        """Tesseractをインストール"""
        from .tesseract_installer import is_tesseract_installed, download_and_install_tesseract

        if is_tesseract_installed():
            messagebox.showinfo("情報", "Tesseractは既にインストールされています")
            self._check_ocr()
            return

        result = messagebox.askyesno(
            "Tesseract OCRのインストール",
            "Tesseract OCRをダウンロードしてインストールします。\n\n"
            "・ダウンロードサイズ: 約70MB\n"
            "・管理者権限が必要な場合があります\n\n"
            "インストールしますか？"
        )

        if not result:
            return

        # インストールボタンを無効化
        self.install_tesseract_btn.config(state='disabled', text="インストール中...")

        def do_install():
            def progress_callback(status, current, total):
                self.root.after(0, lambda: self.ocr_status_label.config(text=f"({status})"))

            success = download_and_install_tesseract(progress_callback)

            def on_complete():
                self.install_tesseract_btn.config(state='normal', text="Tesseractをインストール")
                if success:
                    messagebox.showinfo("完了", "Tesseract OCRのインストールが完了しました")
                    self._check_ocr()
                else:
                    messagebox.showerror(
                        "エラー",
                        "インストールに失敗しました。\n\n"
                        "手動でインストールしてください:\n"
                        "https://github.com/UB-Mannheim/tesseract/wiki"
                    )
                    self._check_ocr()

            self.root.after(0, on_complete)

        import threading
        thread = threading.Thread(target=do_install, daemon=True)
        thread.start()

    def _install_manga_ocr(self):
        """manga-ocrをインストール"""
        result = messagebox.askyesno(
            "manga-ocrのインストール",
            "manga-ocrをインストールします。\n\n"
            "・ダウンロードサイズ: 約2GB（PyTorch含む）\n"
            "・初回実行時にモデル（約400MB）もダウンロードされます\n"
            "・メモリ8GB以上推奨\n\n"
            "インストールしますか？\n"
            "（コマンドプロンプトが開きます）"
        )

        if not result:
            return

        # pipでインストール
        self.install_manga_ocr_btn.config(state='disabled', text="インストール中...")
        self._log("manga-ocrをインストール中... (数分かかります)")

        def do_install():
            import subprocess
            try:
                # pipでmanga-ocrをインストール
                result = subprocess.run(
                    [sys.executable, '-m', 'pip', 'install', 'manga-ocr'],
                    capture_output=True,
                    text=True,
                    timeout=600  # 10分タイムアウト
                )

                def on_complete():
                    self.install_manga_ocr_btn.config(state='normal', text="manga-ocrをインストール")
                    if result.returncode == 0:
                        self._log("manga-ocrのインストールが完了しました")
                        messagebox.showinfo("完了", "manga-ocrのインストールが完了しました。\n\n初回実行時にモデルがダウンロードされます。")
                        self._update_manga_ocr_status()
                    else:
                        self._log(f"インストールエラー: {result.stderr}")
                        messagebox.showerror("エラー", f"インストールに失敗しました:\n{result.stderr[:500]}")

                self.root.after(0, on_complete)

            except subprocess.TimeoutExpired:
                def on_timeout():
                    self.install_manga_ocr_btn.config(state='normal', text="manga-ocrをインストール")
                    self._log("インストールがタイムアウトしました")
                    messagebox.showerror("エラー", "インストールがタイムアウトしました。\n手動でインストールしてください:\npip install manga-ocr")
                self.root.after(0, on_timeout)

            except Exception as e:
                def on_error():
                    self.install_manga_ocr_btn.config(state='normal', text="manga-ocrをインストール")
                    self._log(f"インストールエラー: {str(e)}")
                    messagebox.showerror("エラー", f"インストールに失敗しました:\n{str(e)}")
                self.root.after(0, on_error)

        thread = threading.Thread(target=do_install, daemon=True)
        thread.start()

    def _extract_pdf_text(self):
        """PDFからテキストを直接抽出"""
        from .text_extractor import TextExtractor

        # PDFファイル選択
        pdf_path = filedialog.askopenfilename(
            title="テキストを抽出するPDFを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self.output_folder.get()
        )

        if not pdf_path:
            return

        # テキストが含まれているかチェック
        extractor = TextExtractor()
        has_text = extractor.has_text_content(pdf_path)

        if not has_text:
            result = messagebox.askyesno(
                "確認",
                "このPDFにはテキストが含まれていない可能性があります。\n"
                "（スキャンされた画像のみのPDFかもしれません）\n\n"
                "テキスト抽出を試みますか？\n"
                "（テキストが抽出できない場合は、OCR機能をお試しください）"
            )
            if not result:
                return

        # ボタン無効化
        self.extract_pdf_btn.config(state='disabled')
        self.extract_word_btn.config(state='disabled')
        self.extract_status.config(text="抽出中...", foreground="blue")

        def do_extract():
            try:
                def progress_callback(current, total, status):
                    self.root.after(0, lambda: self.extract_status.config(text=f"{current}/{total}ページ"))

                output_path = extractor.extract_to_file(pdf_path, progress_callback=progress_callback)

                def on_complete():
                    self.extract_pdf_btn.config(state='normal')
                    self.extract_word_btn.config(state='normal')
                    self.extract_status.config(text="完了!", foreground="green")
                    self._log(f"PDFテキスト抽出完了: {output_path}")
                    messagebox.showinfo("完了", f"テキスト抽出が完了しました。\n\n出力ファイル:\n{output_path}")

                self.root.after(0, on_complete)

            except Exception as e:
                def on_error():
                    self.extract_pdf_btn.config(state='normal')
                    self.extract_word_btn.config(state='normal')
                    self.extract_status.config(text="エラー", foreground="red")
                    self._log(f"PDFテキスト抽出エラー: {str(e)}")
                    messagebox.showerror("エラー", f"テキスト抽出中にエラーが発生しました:\n{str(e)}")

                self.root.after(0, on_error)

        thread = threading.Thread(target=do_extract, daemon=True)
        thread.start()

    def _extract_word_text(self):
        """Wordファイルからテキストを直接抽出"""
        from .text_extractor import TextExtractor, check_docx_available

        if not check_docx_available():
            result = messagebox.askyesno(
                "python-docxが必要",
                "Word文書の読み取りにはpython-docxが必要です。\n\n"
                "インストールしますか？"
            )
            if result:
                self._install_python_docx()
            return

        # Wordファイル選択
        word_path = filedialog.askopenfilename(
            title="テキストを抽出するWordファイルを選択",
            filetypes=[
                ("Word files", "*.docx"),
                ("All files", "*.*")
            ],
            initialdir=self.output_folder.get()
        )

        if not word_path:
            return

        # .doc形式のチェック
        if word_path.lower().endswith('.doc') and not word_path.lower().endswith('.docx'):
            messagebox.showerror(
                "非対応形式",
                ".doc形式は直接サポートされていません。\n\n"
                "Wordで開いて「名前を付けて保存」→「.docx」形式で保存してから使用してください。"
            )
            return

        # ボタン無効化
        self.extract_pdf_btn.config(state='disabled')
        self.extract_word_btn.config(state='disabled')
        self.extract_status.config(text="抽出中...", foreground="blue")

        def do_extract():
            try:
                extractor = TextExtractor()

                def progress_callback(current, total, status):
                    self.root.after(0, lambda: self.extract_status.config(text=status))

                output_path = extractor.extract_to_file(word_path, progress_callback=progress_callback)

                def on_complete():
                    self.extract_pdf_btn.config(state='normal')
                    self.extract_word_btn.config(state='normal')
                    self.extract_status.config(text="完了!", foreground="green")
                    self._log(f"Wordテキスト抽出完了: {output_path}")
                    messagebox.showinfo("完了", f"テキスト抽出が完了しました。\n\n出力ファイル:\n{output_path}")

                self.root.after(0, on_complete)

            except Exception as e:
                def on_error():
                    self.extract_pdf_btn.config(state='normal')
                    self.extract_word_btn.config(state='normal')
                    self.extract_status.config(text="エラー", foreground="red")
                    self._log(f"Wordテキスト抽出エラー: {str(e)}")
                    messagebox.showerror("エラー", f"テキスト抽出中にエラーが発生しました:\n{str(e)}")

                self.root.after(0, on_error)

        thread = threading.Thread(target=do_extract, daemon=True)
        thread.start()

    def _install_python_docx(self):
        """python-docxをインストール"""
        self.extract_word_btn.config(state='disabled')
        self.extract_status.config(text="インストール中...", foreground="blue")
        self._log("python-docxをインストール中...")

        def do_install():
            import subprocess
            try:
                result = subprocess.run(
                    [sys.executable, '-m', 'pip', 'install', 'python-docx'],
                    capture_output=True,
                    text=True,
                    timeout=120
                )

                def on_complete():
                    self.extract_word_btn.config(state='normal')
                    if result.returncode == 0:
                        self.extract_status.config(text="インストール完了", foreground="green")
                        self._log("python-docxのインストールが完了しました")
                        messagebox.showinfo("完了", "python-docxのインストールが完了しました。\n\nもう一度「Wordからテキスト抽出」をクリックしてください。")
                    else:
                        self.extract_status.config(text="エラー", foreground="red")
                        self._log(f"インストールエラー: {result.stderr}")
                        messagebox.showerror("エラー", f"インストールに失敗しました:\n{result.stderr[:300]}")

                self.root.after(0, on_complete)

            except Exception as e:
                def on_error():
                    self.extract_word_btn.config(state='normal')
                    self.extract_status.config(text="エラー", foreground="red")
                    self._log(f"インストールエラー: {str(e)}")
                    messagebox.showerror("エラー", f"インストールに失敗しました:\n{str(e)}")

                self.root.after(0, on_error)

        thread = threading.Thread(target=do_install, daemon=True)
        thread.start()

    def _create_ocr_processor(self):
        """OCRプロセッサを作成"""
        from .ocr_processor import OCRProcessor, OCREngine, TextDirection, PreprocessingLevel

        engine = self._get_engine_value()
        ocr_engine = OCREngine.MANGA_OCR if engine == 'manga_ocr' else OCREngine.TESSERACT

        direction_map = {
            'horizontal': TextDirection.HORIZONTAL,
            'vertical': TextDirection.VERTICAL,
            'mixed': TextDirection.MIXED
        }
        text_dir = direction_map.get(self._get_direction_value(), TextDirection.HORIZONTAL)

        preproc_map = {
            'none': PreprocessingLevel.NONE,
            'simple': PreprocessingLevel.SIMPLE,
            'advanced': PreprocessingLevel.ADVANCED
        }
        preproc = preproc_map.get(self._get_preproc_value(), PreprocessingLevel.ADVANCED)

        # 言語設定（テキスト方向に応じて自動設定）
        direction = self._get_direction_value()
        if direction == 'vertical':
            language = 'jpn_vert'
        elif direction == 'mixed':
            language = 'jpn+jpn_vert'
        else:
            language = 'jpn'

        return OCRProcessor(
            language=language,
            engine=ocr_engine,
            text_direction=text_dir,
            preprocessing=preproc
        )

    def _ocr_existing_pdf(self):
        """既存PDFにOCR処理を実行"""
        from .ocr_processor import find_tesseract, check_manga_ocr_available

        engine = self._get_engine_value()

        # エンジンの可用性チェック
        if engine == 'tesseract' and not find_tesseract():
            messagebox.showerror("エラー", "Tesseract OCRがインストールされていません。\n先にTesseractをインストールしてください。")
            return

        if engine == 'manga_ocr' and not check_manga_ocr_available():
            messagebox.showerror("エラー", "manga-ocrがインストールされていません。\n先にmanga-ocrをインストールしてください。")
            return

        # PDFファイル選択
        pdf_path = filedialog.askopenfilename(
            title="OCR処理するPDFを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=self.output_folder.get()
        )

        if not pdf_path:
            return

        # 確認
        engine_name = "manga-ocr（高精度）" if engine == 'manga_ocr' else "Tesseract"
        direction_name = self._direction_display.get()
        result = messagebox.askyesno(
            "確認",
            f"選択したPDFにOCR処理を実行します。\n\n"
            f"ファイル: {os.path.basename(pdf_path)}\n"
            f"エンジン: {engine_name}\n"
            f"本の種類: {direction_name}\n\n"
            f"処理を開始しますか？"
        )

        if not result:
            return

        # ボタン無効化
        self.pdf_ocr_btn.config(state='disabled')
        self.image_ocr_btn.config(state='disabled')
        self.pdf_ocr_status.config(text="処理中...", foreground="blue")

        def do_ocr():
            try:
                ocr = self._create_ocr_processor()

                def progress_callback(current, total, status):
                    self.root.after(0, lambda: self.pdf_ocr_status.config(text=f"{current}/{total}ページ"))

                output_path = ocr.process_pdf_to_file(pdf_path, progress_callback=progress_callback)

                def on_complete():
                    self.pdf_ocr_btn.config(state='normal')
                    self.image_ocr_btn.config(state='normal')
                    self.pdf_ocr_status.config(text="完了!", foreground="green")
                    self._log(f"PDF OCR完了: {output_path}")
                    messagebox.showinfo("完了", f"OCR処理が完了しました。\n\n出力ファイル:\n{output_path}")

                self.root.after(0, on_complete)

            except Exception as e:
                def on_error():
                    self.pdf_ocr_btn.config(state='normal')
                    self.image_ocr_btn.config(state='normal')
                    self.pdf_ocr_status.config(text="エラー", foreground="red")
                    self._log(f"PDF OCRエラー: {str(e)}")
                    messagebox.showerror("エラー", f"OCR処理中にエラーが発生しました:\n{str(e)}")

                self.root.after(0, on_error)

        thread = threading.Thread(target=do_ocr, daemon=True)
        thread.start()

    def _ocr_existing_images(self):
        """既存の画像にOCR処理を実行"""
        from .ocr_processor import find_tesseract, check_manga_ocr_available

        engine = self._get_engine_value()

        # エンジンの可用性チェック
        if engine == 'tesseract' and not find_tesseract():
            messagebox.showerror("エラー", "Tesseract OCRがインストールされていません。\n先にTesseractをインストールしてください。")
            return

        if engine == 'manga_ocr' and not check_manga_ocr_available():
            messagebox.showerror("エラー", "manga-ocrがインストールされていません。\n先にmanga-ocrをインストールしてください。")
            return

        # 画像ファイル選択（複数選択可能）
        image_paths = filedialog.askopenfilenames(
            title="OCR処理する画像を選択（複数選択可）",
            filetypes=[
                ("画像ファイル", "*.png *.jpg *.jpeg *.bmp *.tiff *.gif"),
                ("PNG", "*.png"),
                ("JPEG", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ],
            initialdir=self.output_folder.get()
        )

        if not image_paths:
            return

        # 確認
        engine_name = "manga-ocr（高精度）" if engine == 'manga_ocr' else "Tesseract"
        direction_name = self._direction_display.get()
        result = messagebox.askyesno(
            "確認",
            f"選択した画像にOCR処理を実行します。\n\n"
            f"ファイル数: {len(image_paths)}枚\n"
            f"エンジン: {engine_name}\n"
            f"本の種類: {direction_name}\n\n"
            f"処理を開始しますか？"
        )

        if not result:
            return

        # ボタン無効化
        self.pdf_ocr_btn.config(state='disabled')
        self.image_ocr_btn.config(state='disabled')
        self.pdf_ocr_status.config(text="処理中...", foreground="blue")

        def do_ocr():
            try:
                ocr = self._create_ocr_processor()

                def progress_callback(current, total, status):
                    self.root.after(0, lambda: self.pdf_ocr_status.config(text=f"{current}/{total}枚"))

                results = ocr.process_images(list(image_paths), progress_callback=progress_callback)

                # 出力ファイルパスを決定
                first_image = image_paths[0]
                output_dir = os.path.dirname(first_image)
                output_name = os.path.splitext(os.path.basename(first_image))[0]
                if len(image_paths) > 1:
                    output_name += f"_他{len(image_paths)-1}枚"
                output_path = os.path.join(output_dir, f"{output_name}_ocr.txt")

                ocr.save_ocr_results(results, output_path)

                def on_complete():
                    self.pdf_ocr_btn.config(state='normal')
                    self.image_ocr_btn.config(state='normal')
                    self.pdf_ocr_status.config(text="完了!", foreground="green")
                    self._log(f"画像OCR完了: {output_path}")
                    messagebox.showinfo("完了", f"OCR処理が完了しました。\n\n出力ファイル:\n{output_path}")

                self.root.after(0, on_complete)

            except Exception as e:
                def on_error():
                    self.pdf_ocr_btn.config(state='normal')
                    self.image_ocr_btn.config(state='normal')
                    self.pdf_ocr_status.config(text="エラー", foreground="red")
                    self._log(f"画像OCRエラー: {str(e)}")
                    messagebox.showerror("エラー", f"OCR処理中にエラーが発生しました:\n{str(e)}")

                self.root.after(0, on_error)

        thread = threading.Thread(target=do_ocr, daemon=True)
        thread.start()

    def _toggle_page_input(self):
        if self.stop_mode.get() == 'pages':
            self.pages_entry.config(state='normal')
        else:
            self.pages_entry.config(state='disabled')

    def _update_detect_desc(self, event=None):
        n = self.auto_detect_count.get()
        self.detect_desc_label.config(text=f"(同じ画像が連続{n}回で停止)")

    def _browse_output(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)

    def _select_region(self):
        from .region_selector import RegionSelectorWithPreview
        self.root.iconify()  # 最小化
        self.root.update()

        selector = RegionSelectorWithPreview(self.root)
        region = selector.select_region()

        self.root.deiconify()  # 復帰
        self.root.lift()
        self.root.focus_force()

        if region:
            self.capture_region = region
            self.region_label.config(text=f"選択済: {region[2]-region[0]}x{region[3]-region[1]}px", foreground="green")
            self._log(f"キャプチャ範囲を選択: {region}")
        else:
            self._log("キャプチャ範囲の選択がキャンセルされました")

    def _log(self, message: str):
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def _update_status(self, message: str):
        self.status_label.config(text=message)

    def _start_capture(self):
        if not self.capture_region:
            messagebox.showerror("エラー", "キャプチャ範囲を選択してください")
            return

        if not self.book_title.get().strip():
            messagebox.showerror("エラー", "書籍タイトルを入力してください")
            return

        if self.stop_mode.get() == 'pages':
            try:
                pages = int(self.total_pages.get())
                if pages <= 0:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("エラー", "有効なページ数を入力してください")
                return

        try:
            delay = float(self.delay_time.get())
            if delay < 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("エラー", "有効な待機時間を入力してください")
            return

        result = messagebox.askokcancel(
            "確認",
            "キャプチャを開始します。\n\n"
            "【注意】\n"
            "・Kindleアプリで本を開いた状態にしてください\n"
            "・キャプチャ中はPCを操作しないでください\n"
            "・中断する場合はESCキーを押してください\n\n"
            "3秒後にキャプチャを開始します。"
        )

        if not result:
            return

        self.is_capturing = True
        self.stop_flag = False
        self.start_button.config(state='disabled')
        self.stop_button.config(state='normal')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10), before=self.status_label)
        self.progress_bar.start()

        thread = threading.Thread(target=self._capture_thread, daemon=True)
        thread.start()

    def _stop_capture(self):
        if self.is_capturing:
            self.stop_flag = True
            self._log("中断要求を送信しました...")

    def _capture_thread(self):
        import time
        from .capture import ScreenCapture
        from .pdf_generator import PDFGenerator
        from .privacy_overlay import PrivacyOverlayController

        try:
            for i in range(3, 0, -1):
                if self.stop_flag:
                    self._thread_safe_log("中断されました")
                    return
                self._thread_safe_status(f"キャプチャ開始まで {i} 秒...")
                time.sleep(1)

            # プライバシーモードの設定
            privacy_enabled = self.privacy_mode.get()
            if privacy_enabled:
                self._thread_safe_log("プライバシーモードを有効化...")
                self.privacy_controller = PrivacyOverlayController(self.capture_region, self.root)
                self.privacy_controller.start()
                time.sleep(0.3)

            book_title = self.book_title.get().strip()
            safe_title = "".join(c for c in book_title if c.isalnum() or c in (' ', '-', '_', '.')).strip()
            if not safe_title:
                safe_title = "untitled"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_base = os.path.join(self.output_folder.get(), f"{safe_title}_{timestamp}")
            images_dir = os.path.join(output_base, "images")
            os.makedirs(images_dir, exist_ok=True)

            self._thread_safe_log("キャプチャを開始します...")
            self._thread_safe_status("キャプチャ中...")

            capture = ScreenCapture(
                region=self.capture_region,
                page_direction=self.page_direction.get(),
                delay=float(self.delay_time.get()),
                max_duplicates=int(self.auto_detect_count.get())
            )

            def progress_cb(current, status):
                self._thread_safe_status(f"{status}")

            def check_stop():
                return self.stop_flag

            def on_end_detected():
                # 音を鳴らして通知
                import winsound
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
                self._thread_safe_log("*** 最終ページを検出しました。ESCで停止してください ***")

            stop_mode = self.stop_mode.get()

            # プライバシーモード用コールバック（高速版・画面外移動方式）
            def on_before_capture():
                if self.privacy_controller:
                    self.privacy_controller.hide_for_capture()

            def on_after_capture():
                if self.privacy_controller:
                    self.privacy_controller.show_after_capture()

            if stop_mode == 'pages':
                total = int(self.total_pages.get())
                self.captured_files = capture.capture_all_pages(
                    total_pages=total,
                    output_dir=images_dir,
                    progress_callback=lambda c, t: progress_cb(c, f"{c}/{t}ページ"),
                    check_stop=check_stop,
                    on_before_capture=on_before_capture if privacy_enabled else None,
                    on_after_capture=on_after_capture if privacy_enabled else None
                )
            elif stop_mode == 'manual':
                self.captured_files = capture.capture_until_end(
                    output_dir=images_dir,
                    progress_callback=progress_cb,
                    check_stop=check_stop,
                    manual_mode=True,
                    on_end_detected=on_end_detected,
                    on_before_capture=on_before_capture if privacy_enabled else None,
                    on_after_capture=on_after_capture if privacy_enabled else None
                )
            else:  # auto
                self.captured_files = capture.capture_until_end(
                    output_dir=images_dir,
                    progress_callback=progress_cb,
                    check_stop=check_stop,
                    manual_mode=False,
                    on_before_capture=on_before_capture if privacy_enabled else None,
                    on_after_capture=on_after_capture if privacy_enabled else None
                )

            # プライバシーモードのオーバーレイを削除（PDF生成前）
            if self.privacy_controller:
                cleanup_done = threading.Event()
                def _cleanup():
                    if self.privacy_controller:
                        self.privacy_controller.stop()
                        self.privacy_controller = None
                    cleanup_done.set()
                self.root.after(0, _cleanup)
                cleanup_done.wait(timeout=2)

            if self.stop_flag:
                self._thread_safe_log(f"中断しました（{len(self.captured_files)}ページまでキャプチャ済み）")
            else:
                self._thread_safe_log(f"キャプチャ完了: {len(self.captured_files)}ページ")

            if not self.captured_files:
                self._thread_safe_log("キャプチャされたページがありません")
                return

            # PDF生成
            self._thread_safe_status("PDF生成中...")
            self._thread_safe_log("PDFを生成しています...")

            pdf_generator = PDFGenerator()
            pdf_path = os.path.join(output_base, f"{safe_title}.pdf")

            def pdf_progress(current, total):
                self._thread_safe_status(f"PDF生成中: {current}/{total}")

            pdf_generator.images_to_pdf_direct(self.captured_files, pdf_path, progress_callback=pdf_progress)
            self._thread_safe_log(f"PDF生成完了: {pdf_path}")

            # OCR処理（有効かつ利用可能な場合のみ）
            if self.enable_ocr.get():
                ocr = self._create_ocr_processor()

                if ocr.is_available():
                    engine_name = ocr.get_engine_name()
                    self._thread_safe_status(f"{engine_name}でOCR処理中...")
                    self._thread_safe_log(f"{engine_name}でOCR処理を開始します...")

                    def ocr_progress(current, total, status):
                        self._thread_safe_status(f"OCR: {current}/{total}")

                    ocr_results = ocr.process_images(self.captured_files, progress_callback=ocr_progress)
                    text_path = os.path.join(output_base, f"{safe_title}_ocr.txt")
                    ocr.save_ocr_results(ocr_results, text_path)
                    self._thread_safe_log(f"OCRテキスト保存完了: {text_path}")
                else:
                    engine_name = ocr.get_engine_name()
                    self._thread_safe_log(f"{engine_name}が見つからないため、OCR処理をスキップしました")

            # 完了
            self._thread_safe_status("完了")
            self._thread_safe_log(f"全処理が完了しました。出力先: {output_base}")

            self.root.after(0, lambda: messagebox.showinfo(
                "完了",
                f"処理が完了しました。\n\n"
                f"キャプチャページ数: {len(self.captured_files)}\n"
                f"出力先: {output_base}"
            ))

        except Exception as e:
            self._thread_safe_log(f"エラー: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}"))

        finally:
            # プライバシーオーバーレイの確実なクリーンアップ
            if self.privacy_controller:
                def _cleanup_overlay():
                    if self.privacy_controller:
                        self.privacy_controller.stop()
                        self.privacy_controller = None
                self.root.after(0, _cleanup_overlay)
            self.root.after(0, self._capture_finished)

    def _capture_finished(self):
        self.is_capturing = False
        self.stop_flag = False
        self.start_button.config(state='normal')
        self.stop_button.config(state='disabled')
        self.progress_bar.stop()
        self.progress_bar.pack_forget()  # 非表示に戻す

    def _thread_safe_log(self, message: str):
        self.root.after(0, lambda: self._log(message))

    def _thread_safe_status(self, message: str):
        self.root.after(0, lambda: self._update_status(message))

    def run(self):
        self.root.mainloop()


def main():
    app = MainWindow()
    app.run()


if __name__ == '__main__':
    main()
