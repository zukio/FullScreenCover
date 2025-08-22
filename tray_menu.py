# pystrayで設定メニュー（間隔・表示ファイル選択）
from pystray import Icon, Menu, MenuItem
from PIL import Image
import tkinter as tk
from tkinter import filedialog, simpledialog
import threading
import os
import sys


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def debug_print(*args, **kwargs):
    """デバッグモードが有効な場合のみ出力"""
    # main.pyからDEBUG_MODEをインポートして使用
    try:
        from main import DEBUG_MODE
        if DEBUG_MODE:
            print(*args, **kwargs)
    except ImportError:
        # 単体実行時はすべて出力
        print(*args, **kwargs)


class TrayMenu:
    def __init__(self, controller):
        try:
            debug_print("TrayMenu初期化開始")
            self.controller = controller
            self.icon = Icon('screensaver')
            debug_print("アイコンファイル読み込み中...")
            self.icon.icon = Image.open(get_resource_path('assets/icon.ico'))
            debug_print("メニュー作成中...")
            self.regenerate_menu()
            debug_print("アイコンスレッド起動中...")
            # threading.Thread(target=self.icon.run, daemon=True).start()  # 削除
            debug_print("TrayMenu初期化完了")
        except Exception as e:
            debug_print(f"TrayMenu初期化エラー: {e}")
            import traceback
            traceback.print_exc()

    def increment_interval(self, value):
        current = self.controller.config.get('interval', 300)
        self.set_interval(current + value)

    def decrement_interval(self, value):
        current = self.controller.config.get('interval', 300)
        new_value = max(1, current - value)
        self.set_interval(new_value)

    def toggle_mute_setting(self, icon, item):
        """ミュート設定の切り替え"""
        current = self.controller.config.get('mute_on_screensaver', True)
        self.controller.config['mute_on_screensaver'] = not current
        self.controller.save_config()
        self.regenerate_menu()
        status = "有効" if not current else "無効"
        debug_print(f"スクリーンセーバー時のミュート設定: {status}")

    def regenerate_menu(self):
        """メニューを再生成して現在の設定を反映"""
        mute_enabled = self.controller.config.get('mute_on_screensaver', True)
        mute_text = "🔇 ミュート: 有効" if mute_enabled else "🔊 ミュート: 無効"

        self.icon.menu = Menu(
            MenuItem(
                f'インターバル（秒）: {self.controller.config.get("interval", 300)}',
                Menu(
                    MenuItem('+10秒', lambda: self.increment_interval(10)),
                    MenuItem('-10秒', lambda: self.decrement_interval(10)),
                    MenuItem('30秒', lambda: self.set_interval(30)),
                    MenuItem('60秒', lambda: self.set_interval(60)),
                    MenuItem('120秒', lambda: self.set_interval(120)),
                    MenuItem('300秒', lambda: self.set_interval(300)),
                    MenuItem('600秒', lambda: self.set_interval(600)),
                )
            ),
            MenuItem('画像/動画を選ぶ', self.choose_file),
            MenuItem(mute_text, self.toggle_mute_setting),
            MenuItem('終了', self.on_quit)
        )

    def set_interval(self, sec):
        self.controller.config['interval'] = sec
        self.controller.save_config()
        self.regenerate_menu()

    def choose_file(self, icon, item):
        try:
            # Windows標準のファイルダイアログを使用
            try:
                # win32guiを使ったWindows標準ダイアログ
                import win32gui
                import win32con
                from tkinter import filedialog as fd
                import subprocess
                import os

                # PowerShellを使ってWindows標準のOpenFileDialogを呼び出す
                powershell_cmd = '''
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.InitialDirectory = "{initial_dir}"
$OpenFileDialog.Filter = "画像ファイル (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|動画ファイル (*.mp4;*.avi;*.mov;*.mkv)|*.mp4;*.avi;*.mov;*.mkv|すべてのファイル (*.*)|*.*"
$OpenFileDialog.Title = "スクリーンセーバー用ファイルを選択"
$result = $OpenFileDialog.ShowDialog()
if ($result -eq "OK") {{
    Write-Output $OpenFileDialog.FileName
}}
'''.format(initial_dir=get_resource_path('assets').replace('\\', '\\\\'))

                # PowerShellコマンドを実行
                result = subprocess.run(
                    ['powershell', '-Command', powershell_cmd],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                if result.returncode == 0 and result.stdout.strip():
                    file_path = result.stdout.strip()
                    if os.path.exists(file_path):
                        self.controller.config['media_file'] = file_path
                        self.controller.save_config()
                        debug_print(f"ファイル選択: {file_path}")
                        return

            except Exception as e:
                debug_print(f"PowerShellダイアログエラー: {e}")

            # フォールバック: 従来のtkinterダイアログ（同期実行）
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            root.lift()
            root.focus_force()

            # メインループを一時的に作成
            root.update()

            file_path = filedialog.askopenfilename(
                parent=root,
                title="スクリーンセーバー用ファイルを選択",
                filetypes=[
                    ('画像ファイル', '*.png *.jpg *.jpeg *.bmp *.gif'),
                    ('動画ファイル', '*.mp4 *.avi *.mov *.mkv'),
                    ('すべてのファイル', '*.*')
                ],
                initialdir=get_resource_path('assets')
            )

            if file_path:
                self.controller.config['media_file'] = file_path
                self.controller.save_config()
                debug_print(f"ファイル選択: {file_path}")
            
            root.destroy()
                
        except Exception as e:
            debug_print(f"ファイル選択エラー: {e}")

    def on_quit(self, icon, item):
        self.controller.stop()
        self.icon.stop()

    def run(self):
        """トレイアイコンを実行（ブロッキング）"""
        try:
            if hasattr(self, 'icon'):
                self.icon.run()
        except Exception as e:
            debug_print(f"Tray run error: {e}")

    def stop(self):
        """トレイアイコンを停止"""
        try:
            if hasattr(self, 'icon'):
                self.icon.stop()
        except Exception as e:
            debug_print(f"Tray stop error: {e}")
