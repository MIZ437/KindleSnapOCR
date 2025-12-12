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
        self.root.geometry("600x700")
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
        ocr_frame = ttk.LabelFrame(main_frame, text="OCR設定", padding="10")
        ocr_frame.pack(fill=tk.X, pady=(0, 10))

        ocr_row1 = ttk.Frame(ocr_frame)
        ocr_row1.pack(fill=tk.X)

        self.ocr_check = ttk.Checkbutton(ocr_row1, text="OCR処理を行う", variable=self.enable_ocr)
        self.ocr_check.pack(side=tk.LEFT)

        ttk.Label(ocr_row1, text="言語:").pack(side=tk.LEFT, padx=(20, 0))
        ocr_combo = ttk.Combobox(ocr_row1, textvariable=self.ocr_language, values=['jpn', 'eng', 'jpn+eng'], width=10, state='readonly')
        ocr_combo.pack(side=tk.LEFT, padx=(5, 0))

        self.ocr_status_label = ttk.Label(ocr_row1, text="")
        self.ocr_status_label.pack(side=tk.LEFT, padx=(10, 0))

        self.install_tesseract_btn = ttk.Button(ocr_row1, text="Tesseractをインストール", command=self._install_tesseract)
        self.install_tesseract_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.install_tesseract_btn.pack_forget()  # 初期状態では非表示

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
        from .ocr_processor import find_tesseract
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

    def _toggle_page_input(self):
        if self.stop_mode.get() == 'pages':
            self.pages_entry.config(state='normal')
        else:
            self.pages_entry.config(state='disabled')

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
        from .ocr_processor import OCRProcessor
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

            # プライバシーモード用コールバック（スレッドセーフ・高速版）
            hide_done = threading.Event()
            show_done = threading.Event()

            def on_before_capture():
                if self.privacy_controller and self.privacy_controller.overlay:
                    hide_done.clear()
                    def _hide():
                        try:
                            if self.privacy_controller and self.privacy_controller.overlay:
                                # 透明度を0にして瞬時に非表示（withdrawより高速）
                                self.privacy_controller.overlay.overlay.attributes('-alpha', 0)
                                self.privacy_controller.overlay.overlay.update_idletasks()
                        except:
                            pass
                        hide_done.set()
                    self.root.after(0, _hide)
                    hide_done.wait(timeout=0.5)
                    time.sleep(0.02)  # 最小限の待機

            def on_after_capture():
                if self.privacy_controller and self.privacy_controller.overlay:
                    show_done.clear()
                    def _show():
                        try:
                            if self.privacy_controller and self.privacy_controller.overlay:
                                # 透明度を戻す
                                self.privacy_controller.overlay.overlay.attributes('-alpha', 0.95)
                                self.privacy_controller.overlay.overlay.update_idletasks()
                        except:
                            pass
                        show_done.set()
                    self.root.after(0, _show)
                    show_done.wait(timeout=0.5)

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
                ocr = OCRProcessor(self.ocr_language.get())
                if ocr.is_available():
                    self._thread_safe_status("OCR処理中...")
                    self._thread_safe_log("OCR処理を開始します...")

                    def ocr_progress(current, total, status):
                        self._thread_safe_status(f"OCR: {current}/{total}")

                    ocr_results = ocr.process_images(self.captured_files, progress_callback=ocr_progress)
                    text_path = os.path.join(output_base, f"{safe_title}_ocr.txt")
                    ocr.save_ocr_results(ocr_results, text_path)
                    self._thread_safe_log(f"OCRテキスト保存完了: {text_path}")
                else:
                    self._thread_safe_log("Tesseractが見つからないため、OCR処理をスキップしました")

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
