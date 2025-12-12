"""
画面キャプチャモジュール
指定領域のスクリーンショットを撮影し、ページ送りを自動化する
"""
import time
import os
from PIL import ImageGrab, Image
import pyautogui
import hashlib
from typing import Tuple, Optional, Callable
from pathlib import Path
import keyboard


class ScreenCapture:
    """画面キャプチャとページ送りを管理"""

    def __init__(
        self,
        region: Tuple[int, int, int, int],
        page_direction: str = 'right',
        delay: float = 0.5,
        max_duplicates: int = 10
    ):
        """
        Args:
            region: キャプチャ領域 (left, top, right, bottom)
            page_direction: ページ送り方向 'left' or 'right'
            delay: ページ送り後の待機時間（秒）
            max_duplicates: 自動検出モードでの終了判定回数
        """
        self.region = region
        self.page_direction = page_direction
        self.delay = delay
        self.captured_images = []
        self.last_hash = None
        self.duplicate_count = 0
        self.max_duplicates = max_duplicates
        self.page_count = 0
        self.manual_mode = False  # 手動停止モード
        self.on_end_detected = None  # 終了検出時のコールバック

    def capture_screen(self) -> Image.Image:
        """指定領域のスクリーンショットを撮影"""
        screenshot = ImageGrab.grab(bbox=self.region)
        return screenshot

    def _get_image_hash(self, image: Image.Image) -> str:
        """画像のハッシュ値を計算（重複検出用）"""
        # 小さくリサイズしてハッシュ計算（高速化）
        small = image.resize((32, 32)).convert('L')
        pixels = list(small.getdata())
        return hashlib.md5(bytes(pixels)).hexdigest()

    def is_duplicate(self, image: Image.Image, threshold: float = 0.95) -> bool:
        """前のページと同じかどうかを判定"""
        current_hash = self._get_image_hash(image)

        if self.last_hash is None:
            self.last_hash = current_hash
            return False

        is_same = current_hash == self.last_hash
        self.last_hash = current_hash

        if is_same:
            self.duplicate_count += 1
        else:
            self.duplicate_count = 0

        return is_same

    def turn_page(self):
        """ページを送る"""
        # キャプチャ領域の中央をクリックしてフォーカスを確保
        center_x = (self.region[0] + self.region[2]) // 2
        center_y = (self.region[1] + self.region[3]) // 2
        pyautogui.click(center_x, center_y)
        time.sleep(0.1)

        if self.page_direction == 'left':
            pyautogui.press('left')
        else:
            pyautogui.press('right')

        # 待機中もESCをチェック
        wait_time = self.delay
        check_interval = 0.05
        elapsed = 0
        while elapsed < wait_time:
            if keyboard.is_pressed('escape'):
                return True  # ESC pressed
            time.sleep(check_interval)
            elapsed += check_interval
        return False  # Normal

    def capture_all_pages(
        self,
        total_pages: int,
        output_dir: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        check_stop: Optional[Callable[[], bool]] = None,
        on_before_capture: Optional[Callable[[], None]] = None,
        on_after_capture: Optional[Callable[[], None]] = None
    ) -> list:
        """
        全ページをキャプチャする

        Args:
            total_pages: 総ページ数
            output_dir: 画像保存先ディレクトリ
            progress_callback: 進捗コールバック (current, total)
            check_stop: 停止チェックコールバック
            on_before_capture: キャプチャ前コールバック（プライバシーモード用）
            on_after_capture: キャプチャ後コールバック（プライバシーモード用）

        Returns:
            キャプチャした画像ファイルパスのリスト
        """
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        captured_files = []
        page_num = 1

        # 最初のページをキャプチャ
        time.sleep(0.3)  # ウィンドウフォーカス待ち

        while page_num <= total_pages:
            # 停止チェック
            if check_stop and check_stop():
                break

            # プライバシーモード：キャプチャ前にオーバーレイを非表示
            if on_before_capture:
                on_before_capture()

            # キャプチャ
            image = self.capture_screen()

            # プライバシーモード：キャプチャ後にオーバーレイを再表示
            if on_after_capture:
                on_after_capture()

            # 重複チェック
            if self.is_duplicate(image):
                if self.duplicate_count >= self.max_duplicates:
                    # 最終ページに到達したと判断
                    break
                # 重複の場合はページ送りして再試行
                if self.turn_page():  # ESC pressed
                    break
                continue

            # 画像保存
            filename = f"page_{page_num:04d}.png"
            filepath = os.path.join(output_dir, filename)
            image.save(filepath, 'PNG')
            captured_files.append(filepath)

            # 進捗通知
            if progress_callback:
                progress_callback(page_num, total_pages)

            # ページ送り
            if self.turn_page():  # ESC pressed
                break
            page_num += 1
            self.duplicate_count = 0

        return captured_files

    def capture_until_end(
        self,
        output_dir: str,
        max_pages: int = 1000,
        progress_callback: Optional[Callable[[int, str], None]] = None,
        check_stop: Optional[Callable[[], bool]] = None,
        manual_mode: bool = False,
        on_end_detected: Optional[Callable[[], None]] = None,
        on_before_capture: Optional[Callable[[], None]] = None,
        on_after_capture: Optional[Callable[[], None]] = None
    ) -> list:
        """
        最終ページまで自動でキャプチャする（ページ数不明の場合）

        Args:
            output_dir: 画像保存先ディレクトリ
            max_pages: 最大ページ数（無限ループ防止）
            progress_callback: 進捗コールバック (current, status)
            check_stop: 停止チェックコールバック
            manual_mode: True=手動停止モード（終了検出しても停止せず通知のみ）
            on_end_detected: 終了検出時のコールバック（音・通知用）
            on_before_capture: キャプチャ前コールバック（プライバシーモード用）
            on_after_capture: キャプチャ後コールバック（プライバシーモード用）

        Returns:
            キャプチャした画像ファイルパスのリスト
        """
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        captured_files = []
        page_num = 1
        end_notified = False

        # 最初にKindleウィンドウにフォーカスを当てる
        center_x = (self.region[0] + self.region[2]) // 2
        center_y = (self.region[1] + self.region[3]) // 2
        pyautogui.click(center_x, center_y)
        time.sleep(0.5)

        while page_num <= max_pages:
            if check_stop and check_stop():
                break

            # プライバシーモード：キャプチャ前にオーバーレイを非表示
            if on_before_capture:
                on_before_capture()

            image = self.capture_screen()

            # プライバシーモード：キャプチャ後にオーバーレイを再表示
            if on_after_capture:
                on_after_capture()

            # 最初の3ページは重複検出をスキップ
            if page_num > 3 and self.is_duplicate(image):
                if self.duplicate_count >= self.max_duplicates:
                    if manual_mode:
                        # 手動モード：通知のみ、停止しない
                        if not end_notified and on_end_detected:
                            on_end_detected()
                            end_notified = True
                        if progress_callback:
                            progress_callback(page_num - 1, "最終ページ？ ESCで停止")
                    else:
                        # 自動モード：停止
                        if progress_callback:
                            progress_callback(page_num - 1, "最終ページに到達")
                        break
                if self.turn_page():  # ESC pressed
                    break
                continue

            # 通知フラグをリセット（ページが変わったので）
            end_notified = False

            filename = f"page_{page_num:04d}.png"
            filepath = os.path.join(output_dir, filename)
            image.save(filepath, 'PNG')
            captured_files.append(filepath)

            if progress_callback:
                progress_callback(page_num, f"{page_num}ページ目をキャプチャ")

            if self.turn_page():  # ESC pressed
                break
            page_num += 1
            self.duplicate_count = 0

            # ハッシュを更新
            self.last_hash = self._get_image_hash(image)

        return captured_files


class CaptureConfig:
    """キャプチャ設定を管理"""

    def __init__(self):
        self.region = None
        self.page_direction = 'right'
        self.delay = 0.5
        self.total_pages = None
        self.auto_detect_end = True

    def to_dict(self) -> dict:
        return {
            'region': self.region,
            'page_direction': self.page_direction,
            'delay': self.delay,
            'total_pages': self.total_pages,
            'auto_detect_end': self.auto_detect_end
        }

    def from_dict(self, data: dict):
        self.region = data.get('region')
        self.page_direction = data.get('page_direction', 'right')
        self.delay = data.get('delay', 0.5)
        self.total_pages = data.get('total_pages')
        self.auto_detect_end = data.get('auto_detect_end', True)


if __name__ == '__main__':
    # テスト
    print("3秒後にキャプチャを開始します...")
    time.sleep(3)

    region = (100, 100, 800, 600)
    capture = ScreenCapture(region, 'right', 0.5)
    img = capture.capture_screen()
    img.save('test_capture.png')
    print("キャプチャ完了: test_capture.png")
