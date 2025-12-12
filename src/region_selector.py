"""
範囲選択オーバーレイ
マウスドラッグでキャプチャ領域を指定する
拡大プレビュー機能付き
"""
import tkinter as tk
import ctypes
from PIL import ImageGrab, ImageTk, Image

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    pass


class RegionSelectorWithPreview:
    """画面上で矩形領域を選択（拡大プレビュー付き）"""

    def __init__(self, parent=None):
        self.parent = parent
        self.start_x = 0
        self.start_y = 0
        self.rect = None
        self.selected_region = None
        self.screenshot = None  # 事前に撮影したスクリーンショット
        self.preview_size = 200  # プレビューウィンドウのサイズ
        self.zoom_factor = 3    # 拡大率
        self.preview_label = None
        self.crosshair_h = None
        self.crosshair_v = None

    def select_region(self):
        """
        オーバーレイを表示してユーザーに範囲選択させる
        Returns: (left, top, right, bottom) or None if cancelled
        """
        # オーバーレイ表示前にスクリーンショットを撮る
        self.screenshot = ImageGrab.grab()

        # 親がある場合はToplevel、ない場合はTk
        if self.parent:
            self.root = tk.Toplevel(self.parent)
        else:
            self.root = tk.Tk()

        self.root.attributes('-fullscreen', True)
        self.root.attributes('-alpha', 0.3)
        self.root.attributes('-topmost', True)
        self.root.configure(bg='black')
        self.root.config(cursor='cross')

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        self.canvas = tk.Canvas(
            self.root,
            width=screen_width,
            height=screen_height,
            bg='black',
            highlightthickness=0
        )
        self.canvas.pack()

        # 説明テキスト
        self.canvas.create_text(
            screen_width // 2, 30,
            text="ドラッグで範囲選択 | ESCでキャンセル",
            font=('Arial', 16, 'bold'),
            fill='white'
        )
        self.canvas.create_text(
            screen_width // 2, 55,
            text="右下に拡大プレビューが表示されます",
            font=('Arial', 12),
            fill='#aaaaaa'
        )

        # 拡大プレビュー用のフレーム（右下に配置）
        preview_frame = tk.Frame(self.root, bg='#333333', bd=2, relief='solid')
        preview_x = screen_width - self.preview_size - 30
        preview_y = screen_height - self.preview_size - 80
        preview_frame.place(x=preview_x, y=preview_y)

        # プレビューラベル
        self.preview_label = tk.Label(
            preview_frame,
            width=self.preview_size,
            height=self.preview_size,
            bg='#222222'
        )
        self.preview_label.pack()

        # 座標表示ラベル
        self.coord_label = tk.Label(
            preview_frame,
            text="X: 0, Y: 0",
            font=('Consolas', 10),
            fg='white',
            bg='#333333'
        )
        self.coord_label.pack(pady=2)

        # イベントバインド
        self.canvas.bind('<Button-1>', self._on_press)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        self.canvas.bind('<Motion>', self._on_motion)
        self.root.bind('<Escape>', self._on_cancel)

        self.root.grab_set()
        self.root.focus_force()

        if self.parent:
            self.root.wait_window()
        else:
            self.root.mainloop()

        return self.selected_region

    def _update_preview(self, x, y):
        """マウス位置周辺の拡大プレビューを更新"""
        if not self.screenshot or not self.preview_label:
            return

        try:
            # 拡大する領域のサイズ（プレビューサイズ / 拡大率）
            capture_size = self.preview_size // self.zoom_factor

            # キャプチャ領域を計算（マウス位置を中心に）
            half = capture_size // 2
            left = max(0, x - half)
            top = max(0, y - half)
            right = min(self.screenshot.width, x + half)
            bottom = min(self.screenshot.height, y + half)

            # スクリーンショットから切り出し
            cropped = self.screenshot.crop((left, top, right, bottom))

            # 拡大（nearest neighborで鮮明に）
            zoomed = cropped.resize(
                (self.preview_size, self.preview_size),
                Image.Resampling.NEAREST
            )

            # 中央に十字線を描画
            from PIL import ImageDraw
            draw = ImageDraw.Draw(zoomed)
            center = self.preview_size // 2
            line_color = '#ff0000'
            # 横線
            draw.line([(0, center), (self.preview_size, center)], fill=line_color, width=1)
            # 縦線
            draw.line([(center, 0), (center, self.preview_size)], fill=line_color, width=1)

            # tkinter用に変換
            self.preview_image = ImageTk.PhotoImage(zoomed)
            self.preview_label.configure(image=self.preview_image)

            # 座標を更新
            self.coord_label.configure(text=f"X: {x}, Y: {y}")

        except Exception:
            pass

    def _on_motion(self, event):
        """マウス移動時にプレビューを更新"""
        self._update_preview(event.x, event.y)

    def _on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y,
            self.start_x, self.start_y,
            outline='red', width=3
        )

    def _on_drag(self, event):
        if self.rect:
            self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)
        self._update_preview(event.x, event.y)

    def _on_release(self, event):
        left = min(self.start_x, event.x)
        top = min(self.start_y, event.y)
        right = max(self.start_x, event.x)
        bottom = max(self.start_y, event.y)

        if right - left > 50 and bottom - top > 50:
            self.selected_region = (left, top, right, bottom)
            self.root.destroy()
        else:
            if self.rect:
                self.canvas.delete(self.rect)
                self.rect = None

    def _on_cancel(self, event):
        self.selected_region = None
        self.root.destroy()
