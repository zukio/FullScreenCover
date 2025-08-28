"""
マウスカーソー表示制御モジュール
WindowsAPIを使用してシステムレベルでカーソーの表示/非表示を制御
"""

import ctypes
from ctypes import wintypes
import threading


class CursorController:
    """マウスカーソーの表示/非表示を制御するクラス"""

    def __init__(self):
        self._cursor_hidden = False
        self._lock = threading.Lock()
        self._original_cursor_count = 0

        # Windows API関数の定義
        self._user32 = ctypes.windll.user32
        self._ShowCursor = self._user32.ShowCursor
        self._ShowCursor.argtypes = [wintypes.BOOL]
        self._ShowCursor.restype = ctypes.c_int

    def hide_cursor(self):
        """マウスカーソーを非表示にする"""
        with self._lock:
            if not self._cursor_hidden:
                try:
                    # カーソーが完全に非表示になるまで繰り返し呼び出し
                    count = 0
                    while count >= 0 and count < 100:  # 無限ループ防止
                        count = self._ShowCursor(False)

                    self._cursor_hidden = True
                    self._original_cursor_count = count
                    print(f"マウスカーソー非表示完了 (count: {count})")
                    return True
                except Exception as e:
                    print(f"マウスカーソー非表示エラー: {e}")
                    return False
            return True

    def show_cursor(self):
        """マウスカーソーを表示する"""
        with self._lock:
            if self._cursor_hidden:
                try:
                    # カーソーが表示されるまで繰り返し呼び出し
                    count = self._original_cursor_count
                    while count < 0:
                        count = self._ShowCursor(True)

                    self._cursor_hidden = False
                    print(f"マウスカーソー表示復元完了 (count: {count})")
                    return True
                except Exception as e:
                    print(f"マウスカーソー表示復元エラー: {e}")
                    return False
            return True

    def is_cursor_hidden(self):
        """カーソーが非表示状態かどうかを返す"""
        return self._cursor_hidden


# グローバルカーソーコントローラーのインスタンス
_cursor_controller = None
_cursor_controller_lock = threading.Lock()


def get_cursor_controller():
    """CursorControllerのシングルトンインスタンスを取得"""
    global _cursor_controller

    if _cursor_controller is None:
        with _cursor_controller_lock:
            if _cursor_controller is None:
                _cursor_controller = CursorController()

    return _cursor_controller


def hide_system_cursor():
    """システムレベルでマウスカーソーを非表示（便利関数）"""
    return get_cursor_controller().hide_cursor()


def show_system_cursor():
    """システムレベルでマウスカーソーを表示（便利関数）"""
    return get_cursor_controller().show_cursor()


def is_system_cursor_hidden():
    """システムカーソーが非表示状態かどうかを確認（便利関数）"""
    return get_cursor_controller().is_cursor_hidden()


if __name__ == "__main__":
    # テスト用コード
    import time

    print("=== CursorController テスト ===")
    controller = get_cursor_controller()

    print("3秒後にカーソーを非表示にします...")
    time.sleep(3)

    controller.hide_cursor()
    print("カーソーが非表示になりました。3秒間待機...")
    time.sleep(3)

    controller.show_cursor()
    print("カーソーが表示されました。")
