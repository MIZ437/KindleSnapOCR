"""
KindleSnapOCR - Kindle本PDF化ツール
メインエントリーポイント
"""
import sys
import os

# srcディレクトリをパスに追加
if getattr(sys, 'frozen', False):
    # PyInstallerでビルドされた場合
    application_path = os.path.dirname(sys.executable)
else:
    # 通常のPython実行
    application_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, application_path)
sys.path.insert(0, os.path.join(application_path, 'src'))


def main():
    """アプリケーションのメインエントリーポイント"""
    from src.gui import MainWindow

    app = MainWindow()
    app.run()


if __name__ == '__main__':
    main()
