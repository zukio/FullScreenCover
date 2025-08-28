"""
ディスプレイ関連のユーティリティモジュール
複数ディスプレイの情報取得と制御を行う
"""
import tkinter as tk
import threading
from typing import List, Dict, Optional, Tuple


class DisplayInfo:
    """ディスプレイ情報を格納するクラス"""

    def __init__(self, index: int, x: int, y: int, width: int, height: int, is_primary: bool = False):
        self.index = index
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.is_primary = is_primary

    def __str__(self):
        primary_text = " (プライマリ)" if self.is_primary else ""
        return f"ディスプレイ {self.index + 1}: {self.width}x{self.height}{primary_text}"

    def get_geometry(self) -> str:
        """tkinterのgeometry文字列を生成"""
        return f"{self.width}x{self.height}+{self.x}+{self.y}"


class DisplayManager:
    """ディスプレイ管理クラス"""

    def __init__(self):
        self._displays: List[DisplayInfo] = []
        self._refresh_displays()

    def _refresh_displays(self):
        """ディスプレイ情報を更新"""
        self._displays.clear()

        try:
            # 一時的なtkinterウィンドウを作成してディスプレイ情報を取得
            temp_root = tk.Tk()
            temp_root.withdraw()  # ウィンドウを非表示にする

            try:
                # プライマリディスプレイの情報を取得
                screen_width = temp_root.winfo_screenwidth()
                screen_height = temp_root.winfo_screenheight()

                # プライマリディスプレイを追加
                primary_display = DisplayInfo(
                    0, 0, 0, screen_width, screen_height, True)
                self._displays.append(primary_display)

                # 追加のディスプレイ情報を取得を試行
                # （この部分は環境によって異なる可能性があります）
                try:
                    # Windowsの場合、追加のディスプレイ情報を取得
                    import win32api
                    import win32con

                    monitors = win32api.EnumDisplayMonitors()
                    for i, (hmon, hdc, rect) in enumerate(monitors):
                        if i == 0:
                            continue  # プライマリディスプレイはすでに追加済み

                        x, y, right, bottom = rect
                        width = right - x
                        height = bottom - y

                        display = DisplayInfo(i, x, y, width, height, False)
                        self._displays.append(display)

                except ImportError:
                    # win32apiが利用できない場合は、プライマリディスプレイのみ
                    print("win32api not available, using primary display only")

            finally:
                temp_root.destroy()

        except Exception as e:
            print(f"ディスプレイ情報取得エラー: {e}")
            # フォールバック: デフォルトのディスプレイ情報
            self._displays.append(DisplayInfo(0, 0, 0, 1920, 1080, True))

    def get_displays(self) -> List[DisplayInfo]:
        """利用可能なディスプレイのリストを取得"""
        return self._displays.copy()

    def get_display_count(self) -> int:
        """ディスプレイ数を取得"""
        return len(self._displays)

    def get_display_by_index(self, index: int) -> Optional[DisplayInfo]:
        """インデックスでディスプレイ情報を取得"""
        if 0 <= index < len(self._displays):
            return self._displays[index]
        return None

    def get_primary_display(self) -> DisplayInfo:
        """プライマリディスプレイの情報を取得"""
        for display in self._displays:
            if display.is_primary:
                return display
        # プライマリが見つからない場合は最初のディスプレイを返す
        return self._displays[0] if self._displays else DisplayInfo(0, 0, 0, 1920, 1080, True)

    def refresh(self):
        """ディスプレイ情報を再取得"""
        self._refresh_displays()

    def get_display_names(self) -> List[str]:
        """ディスプレイ名のリストを取得（選択用）"""
        return [str(display) for display in self._displays]


# グローバルインスタンス
_display_manager = None
_display_manager_lock = threading.Lock()


def get_display_manager() -> DisplayManager:
    """DisplayManagerのシングルトンインスタンスを取得"""
    global _display_manager

    if _display_manager is None:
        with _display_manager_lock:
            if _display_manager is None:
                _display_manager = DisplayManager()

    return _display_manager


def get_available_displays() -> List[DisplayInfo]:
    """利用可能なディスプレイのリストを取得（便利関数）"""
    return get_display_manager().get_displays()


def get_display_by_index(index: int) -> Optional[DisplayInfo]:
    """インデックスでディスプレイ情報を取得（便利関数）"""
    return get_display_manager().get_display_by_index(index)


def get_primary_display() -> DisplayInfo:
    """プライマリディスプレイの情報を取得（便利関数）"""
    return get_display_manager().get_primary_display()


if __name__ == "__main__":
    # テスト用コード
    manager = get_display_manager()
    displays = manager.get_displays()

    print(f"検出されたディスプレイ数: {len(displays)}")
    for display in displays:
        print(f"  {display}")
        print(f"    位置: ({display.x}, {display.y})")
        print(f"    サイズ: {display.width}x{display.height}")
        print(f"    Geometry: {display.get_geometry()}")
