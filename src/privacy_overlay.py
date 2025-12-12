"""
プライバシーオーバーレイモジュール
キャプチャ中に画面を隠すための黒いオーバーレイウィンドウ
Windows APIを使用した高速制御
"""
import tkinter as tk
import threading
import ctypes
from ctypes import wintypes
import time


# Windows API定数
GWL_EXSTYLE = -20
WS_EX_LAYERED = 0x80000
WS_EX_TRANSPARENT = 0x20
WS_EX_TOOLWINDOW = 0x80
WS_EX_TOPMOST = 0x8
LWA_ALPHA = 0x2

# Windows API関数
user32 = ctypes.windll.user32
SetWindowLongW = user32.SetWindowLongW
GetWindowLongW = user32.GetWindowLongW
SetLayeredWindowAttributes = user32.SetLayeredWindowAttributes
SetWindowPos = user32.SetWindowPos

# SetWindowPos flags
HWND_TOPMOST = -1
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_SHOWWINDOW = 0x0040
SWP_NOACTIVATE = 0x0010


class FastPrivacyOverlay:
    """高速プライバシーオーバーレイ（Windows API使用）"""

    def __init__(self, region, parent=None):
        """
        Args:
            region: (left, top, right, bottom) キャプチャ領域
            parent: 親ウィンドウ
        """
        self.region = region
        self.parent = parent
        self.overlay = None
        self.hwnd = None
        self.is_visible = True
        self._lock = threading.Lock()
        # 画面外の位置（非表示用）
        self.hidden_pos = (-10000, -10000)
        self.normal_pos = (region[0], region[1])

    def create(self):
        """オーバーレイウィンドウを作成"""
        if self.overlay is not None:
            return

        left, top, right, bottom = self.region
        width = right - left
        height = bottom - top

        if self.parent:
            self.overlay = tk.Toplevel(self.parent)
        else:
            self.overlay = tk.Tk()

        self.overlay.overrideredirect(True)
        self.overlay.geometry(f"{width}x{height}+{left}+{top}")
        self.overlay.configure(bg='black')
        self.overlay.attributes('-topmost', True)

        # メッセージ表示
        frame = tk.Frame(self.overlay, bg='black')
        frame.place(relx=0.5, rely=0.5, anchor='center')

        tk.Label(
            frame,
            text="🔒 Privacy Mode",
            font=('Segoe UI', 16, 'bold'),
            fg='#444444',
            bg='black'
        ).pack()

        tk.Label(
            frame,
            text="キャプチャ中...",
            font=('Segoe UI', 11),
            fg='#333333',
            bg='black'
        ).pack(pady=(10, 0))

        tk.Label(
            frame,
            text="ESCで中断",
            font=('Segoe UI', 10),
            fg='#555555',
            bg='black'
        ).pack(pady=(5, 0))

        self.overlay.update()

        # ウィンドウハンドルを取得
        self.hwnd = ctypes.windll.user32.GetParent(self.overlay.winfo_id())

        # レイヤードウィンドウに設定
        style = GetWindowLongW(self.hwnd, GWL_EXSTYLE)
        new_style = style | WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW
        SetWindowLongW(self.hwnd, GWL_EXSTYLE, new_style)

        # 初期透明度を設定（255 = 完全不透明）
        SetLayeredWindowAttributes(self.hwnd, 0, 250, LWA_ALPHA)

        # 最前面に設定
        SetWindowPos(self.hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE)

        self.is_visible = True
        self.overlay.update()

    def set_alpha(self, alpha):
        """透明度を設定（0-255）- Windows API直接呼び出しで高速"""
        if self.hwnd:
            SetLayeredWindowAttributes(self.hwnd, 0, alpha, LWA_ALPHA)

    def hide_instant(self):
        """瞬時に非表示（画面外に移動）"""
        with self._lock:
            if self.overlay and self.is_visible:
                try:
                    # 画面外に移動（最も高速な非表示方法）
                    self.overlay.geometry(f"+{self.hidden_pos[0]}+{self.hidden_pos[1]}")
                    self.is_visible = False
                except tk.TclError:
                    pass

    def show_instant(self):
        """瞬時に表示（元の位置に戻す）"""
        with self._lock:
            if self.overlay and not self.is_visible:
                try:
                    left, top, right, bottom = self.region
                    width = right - left
                    height = bottom - top
                    self.overlay.geometry(f"{width}x{height}+{left}+{top}")
                    # 最前面を確保
                    if self.hwnd:
                        SetWindowPos(self.hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                                     SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE)
                    self.is_visible = True
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
                self.hwnd = None
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
        self._action_done = threading.Event()

    def start(self):
        """オーバーレイを作成して表示"""
        def _create():
            self.overlay = FastPrivacyOverlay(self.region, self.root)
            self.overlay.create()
            self._created.set()

        self.root.after(0, _create)
        self._created.wait(timeout=5)

    def hide_for_capture(self):
        """キャプチャ用に瞬時に非表示"""
        if not self.overlay:
            return

        self._action_done.clear()

        def _hide():
            if self.overlay:
                self.overlay.hide_instant()
            self._action_done.set()

        self.root.after(0, _hide)
        self._action_done.wait(timeout=1)
        # 非表示が反映されるまでの最小待機
        time.sleep(0.01)

    def show_after_capture(self):
        """キャプチャ後に瞬時に表示"""
        if not self.overlay:
            return

        self._action_done.clear()

        def _show():
            if self.overlay:
                self.overlay.show_instant()
            self._action_done.set()

        self.root.after(0, _show)
        self._action_done.wait(timeout=1)

    def stop(self):
        """オーバーレイを削除"""
        def _destroy():
            if self.overlay:
                self.overlay.destroy()
                self.overlay = None

        self.root.after(0, _destroy)


# 後方互換性のためのエイリアス
PrivacyOverlay = FastPrivacyOverlay
