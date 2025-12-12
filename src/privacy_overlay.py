"""
プライバシーオーバーレイモジュール
キャプチャ中に画面を隠すための黒いオーバーレイウィンドウ
クリック透過機能付き
"""
import tkinter as tk
import threading
import ctypes
import time


# Windows API定数
GWL_EXSTYLE = -20
WS_EX_LAYERED = 0x80000
WS_EX_TRANSPARENT = 0x20
WS_EX_TOOLWINDOW = 0x80


def make_click_through(window):
    """ウィンドウをクリック透過にする（Windowsのみ）"""
    try:
        # tkinterウィンドウのハンドルを取得
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())

        # 現在の拡張スタイルを取得
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)

        # クリック透過とレイヤードウィンドウを追加
        # WS_EX_TOOLWINDOW: タスクバーに表示しない
        new_style = style | WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_style)

        return True
    except Exception:
        return False


class PrivacyOverlay:
    """キャプチャ領域を覆う黒いオーバーレイ（クリック透過対応）"""

    def __init__(self, region, parent=None):
        """
        Args:
            region: (left, top, right, bottom) キャプチャ領域
            parent: 親ウィンドウ（Tkのルートまたはトップレベル）
        """
        self.region = region
        self.parent = parent
        self.overlay = None
        self.is_visible = False
        self._lock = threading.Lock()
        self.hwnd = None

    def create(self):
        """オーバーレイウィンドウを作成（クリック透過）"""
        if self.overlay is not None:
            return

        left, top, right, bottom = self.region
        width = right - left
        height = bottom - top

        if self.parent:
            self.overlay = tk.Toplevel(self.parent)
        else:
            self.overlay = tk.Tk()

        self.overlay.overrideredirect(True)  # 枠なし
        self.overlay.geometry(f"{width}x{height}+{left}+{top}")
        self.overlay.configure(bg='black')
        self.overlay.attributes('-topmost', True)
        self.overlay.attributes('-alpha', 0.95)  # 少しだけ透明に

        # メッセージ表示
        label = tk.Label(
            self.overlay,
            text="Privacy Mode\n\nキャプチャ中...\nESCで中断",
            font=('Arial', 14, 'bold'),
            fg='#888888',  # グレーで目立たなく
            bg='black'
        )
        label.place(relx=0.5, rely=0.5, anchor='center')

        self.overlay.update()

        # クリック透過を設定（マウスクリックがオーバーレイを通過する）
        make_click_through(self.overlay)

        self.is_visible = True
        self.overlay.update()

    def show(self):
        """オーバーレイを表示"""
        with self._lock:
            if self.overlay and not self.is_visible:
                try:
                    self.overlay.deiconify()
                    self.overlay.attributes('-topmost', True)
                    self.overlay.update_idletasks()
                    self.is_visible = True
                except tk.TclError:
                    pass

    def hide(self):
        """オーバーレイを一時的に非表示"""
        with self._lock:
            if self.overlay and self.is_visible:
                try:
                    self.overlay.withdraw()
                    self.overlay.update_idletasks()
                    self.is_visible = False
                except tk.TclError:
                    pass

    def destroy(self):
        """オーバーレイを完全に削除"""
        with self._lock:
            if self.overlay:
                try:
                    self.overlay.destroy()
                except tk.TclError:
                    pass
                self.overlay = None
                self.is_visible = False


class PrivacyOverlayController:
    """別スレッドからオーバーレイを制御するためのコントローラー"""

    def __init__(self, region, root):
        """
        Args:
            region: キャプチャ領域
            root: tkinterのルートウィンドウ
        """
        self.region = region
        self.root = root
        self.overlay = None
        self._created = threading.Event()

    def start(self):
        """オーバーレイを作成して表示（メインスレッドで実行）"""
        def _create():
            self.overlay = PrivacyOverlay(self.region, self.root)
            self.overlay.create()
            self._created.set()

        self.root.after(0, _create)
        self._created.wait(timeout=5)

    def stop(self):
        """オーバーレイを削除（メインスレッドで実行）"""
        def _destroy():
            if self.overlay:
                self.overlay.destroy()
                self.overlay = None

        self.root.after(0, _destroy)
