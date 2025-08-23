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

    def toggle_video_suppress_setting(self, icon, item):
        """動画抑制設定の切り替え"""
        current = self.controller.config.get('suppress_during_video', True)
        self.controller.config['suppress_during_video'] = not current
        self.controller.save_config()
        self.regenerate_menu()
        status = "有効" if not current else "無効"
        debug_print(f"動画再生中の抑制設定: {status}")

    def regenerate_menu(self):
        """メニューを再生成して現在の設定を反映"""
        mute_enabled = self.controller.config.get('mute_on_screensaver', True)
        mute_text = "🔇 ミュート: ☑ 有効" if mute_enabled else "🔊 ミュート: ☐ 無効"

        video_suppress_enabled = self.controller.config.get(
            'suppress_during_video', True)
        video_suppress_text = "🎬 動画再生中は抑制: ☑ 有効" if video_suppress_enabled else "🎬 動画再生中は抑制: ☐ 無効"

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
            MenuItem(video_suppress_text, self.toggle_video_suppress_setting),
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
}} else {{
    Write-Output "CANCELLED"
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
                    output = result.stdout.strip()
                    if output != "CANCELLED" and os.path.exists(output):
                        self.controller.config['media_file'] = output
                        self.controller.save_config()
                        debug_print(f"ファイル選択: {output}")
                        return
                    elif output == "CANCELLED":
                        debug_print("ファイル選択がキャンセルされました")
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

            # ファイルが選択された場合のみ保存
            if file_path and file_path.strip():
                self.controller.config['media_file'] = file_path
                self.controller.save_config()
                debug_print(f"ファイル選択: {file_path}")
            else:
                debug_print("ファイル選択がキャンセルされました")

            root.destroy()

        except Exception as e:
            debug_print(f"ファイル選択エラー: {e}")
            # エラーが発生しても致命的エラーとして扱わない

    def on_quit(self, icon, item):
        """アプリケーション終了処理"""
        debug_print("終了メニューが選択されました")
        try:
            # コントローラーを先に停止
            if hasattr(self, 'controller') and self.controller:
                self.controller.stop()

            # トレイアイコンを停止
            if hasattr(self, 'icon') and self.icon:
                self.icon.stop()
        except Exception as e:
            debug_print(f"終了処理エラー: {e}")
            # 強制終了
            import sys
            sys.exit(0)

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
            debug_print("TrayMenu停止処理開始")
            if hasattr(self, 'icon') and self.icon:
                self.icon.stop()
            debug_print("TrayMenu停止処理完了")
        except Exception as e:
            debug_print(f"Tray stop error: {e}")
            # 強制的にスレッドを終了
            import sys
            import os
            os._exit(0)
