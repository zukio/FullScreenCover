# pystrayで設定メニュー（間隔・表示ファイル選択）
from pystray import Icon, Menu, MenuItem
from PIL import Image
import tkinter as tk
from tkinter import filedialog, simpledialog
import threading
import os
import sys
from modules.utils.display_utils import get_display_manager


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
            self.is_paused = False  # 一時停止状態を管理する変数
            self.display_manager = get_display_manager()  # ディスプレイマネージャーを取得
            debug_print("アイコンファイル読み込み中...")
            self.icon.icon = Image.open(get_resource_path('assets/icon.ico'))
            debug_print("メニュー作成中...")
            self.regenerate_menu()
            debug_print("アイコンスレッド起動中...")
            debug_print("TrayMenu初期化完了")
        except Exception as e:
            debug_print(f"TrayMenu初期化エラー: {e}")
            import traceback
            traceback.print_exc()

    def increment_interval(self, value):
        def _increment(icon, item):
            current = self.controller.config.get('interval', 300)
            self.set_interval_value(current + value)
        return _increment

    def decrement_interval(self, value):
        def _decrement(icon, item):
            current = self.controller.config.get('interval', 300)
            new_value = max(1, current - value)
            self.set_interval_value(new_value)
        return _decrement

    def set_interval(self, sec):
        def _set_interval(icon, item):
            self.set_interval_value(sec)
        return _set_interval

    def set_interval_value(self, sec):
        """実際にインターバル値を設定するメソッド"""
        self.controller.config['interval'] = sec
        self.controller.save_config()
        self.regenerate_menu()

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

    def toggle_presentation_mode_setting(self, icon, item):
        """プレゼンテーションモード設定の切り替え（UI上では一括設定として動作）"""
        current = self.controller.config.get('enable_presentation_mode', False)
        self.controller.config['enable_presentation_mode'] = not current
        self.controller.save_config()
        self.regenerate_menu()
        status = "有効" if not current else "無効"
        debug_print(f"プレゼンテーションモード設定: {status}")

        # 有効になった場合、設定されている機能の詳細をログ出力
        if not current:  # 有効になった場合
            features = self.controller.config.get('presentation_features', {})
            enabled_features = []
            if features.get('disable_screensaver', True):
                enabled_features.append("スクリーンセーバー無効化")
            if features.get('prevent_sleep', True):
                enabled_features.append("スリープ防止")
            if features.get('block_notifications', False):
                enabled_features.append("通知ブロック")

            if enabled_features:
                debug_print(f"  有効な機能: {', '.join(enabled_features)}")
            debug_print(
                "  ※ 高度な機能設定はconfig.jsonの'presentation_features'で個別制御可能")

    def toggle_pause(self, icon, item):
        """一時停止/再開の切り替え"""
        self.is_paused = not self.is_paused
        status = "一時停止中" if self.is_paused else "再開中"
        debug_print(f"アプリの状態: {status}")

        # 一時停止中はコントローラーの動作を停止
        if self.is_paused:
            self.controller.pause()  # pause メソッドを呼び出す
        else:
            self.controller.resume()  # resume メソッドを呼び出す

        self.regenerate_menu()

    def set_display_mode(self, mode):
        """ディスプレイ表示モードを設定"""
        def _set_mode(icon, item):
            self.controller.config['display_mode'] = mode
            if mode != 'specific':
                # 特定ディスプレイ以外の場合はdisplay_indexをクリア
                self.controller.config['display_index'] = None
            self.controller.save_config()
            self.regenerate_menu()
            debug_print(f"ディスプレイ表示モード: {mode}")
        return _set_mode

    def set_specific_display(self, display_index):
        """特定のディスプレイを選択"""
        def _set_display(icon, item):
            self.controller.config['display_mode'] = 'specific'
            self.controller.config['display_index'] = display_index
            self.controller.save_config()
            self.regenerate_menu()
            debug_print(f"表示ディスプレイ: {display_index + 1}")
        return _set_display

    def get_display_menu_items(self):
        """ディスプレイ選択メニュー項目を生成"""
        current_mode = self.controller.config.get('display_mode', 'primary')
        current_index = self.controller.config.get('display_index', None)

        items = []

        # 全ディスプレイ
        check_all = "☑" if current_mode == 'all' else "☐"
        items.append(MenuItem(f'{check_all} 全ディスプレイ',
                     self.set_display_mode('all')))

        # 各ディスプレイ
        displays = self.display_manager.get_displays()
        if len(displays) > 0:
            items.append(MenuItem('-', None))  # セパレータ
            for i, display in enumerate(displays):
                # プライマリディスプレイが選択されている場合の判定を調整
                is_selected = (current_mode == 'specific' and current_index == i) or \
                              (current_mode == 'primary' and display.is_primary)
                check_specific = "☑" if is_selected else "☐"
                display_name = f"ディスプレイ {i + 1}"
                if display.is_primary:
                    display_name += " (プライマリ)"
                display_name += f" - {display.width}x{display.height}"
                items.append(MenuItem(
                    f'{check_specific} {display_name}',
                    self.set_specific_display(i)))

        return items

    def regenerate_menu(self):
        """メニューを再生成して現在の設定を反映"""
        mute_enabled = self.controller.config.get('mute_on_screensaver', True)
        mute_text = "☑ 遮蔽時はミュートする" if mute_enabled else "☐ 遮蔽時はミュートする（していません）"

        video_suppress_enabled = self.controller.config.get(
            'suppress_during_video', True)
        video_suppress_text = "☑ 動画再生中は待機" if video_suppress_enabled else "☐ 動画再生中は待機（していません）"

        presentation_enabled = self.controller.config.get(
            'enable_presentation_mode', True)
        presentation_enabled_text = "☑ 通知やスリープをブロック" if presentation_enabled else "☐ 通知やスリープをブロック（していません）"

        pause_text = "⏸ 一時停止" if not self.is_paused else "▶ 再開（停止中）"

        # ディスプレイ設定の表示テキスト
        display_mode = self.controller.config.get('display_mode', 'primary')
        display_index = self.controller.config.get('display_index', None)

        if display_mode == 'primary':
            # プライマリディスプレイの場合、プライマリディスプレイの番号を表示
            primary_display = self.display_manager.get_primary_display()
            if primary_display:
                display_text = f"ディスプレイ {primary_display.index + 1} (プライマリ)"
            else:
                display_text = "プライマリディスプレイ"
        elif display_mode == 'all':
            display_text = "全ディスプレイ"
        elif display_mode == 'specific' and display_index is not None:
            target_display = self.display_manager.get_display_by_index(
                display_index)
            if target_display:
                primary_text = " (プライマリ)" if target_display.is_primary else ""
                display_text = f"ディスプレイ {display_index + 1}{primary_text}"
            else:
                display_text = f"ディスプレイ {display_index + 1}"
        else:
            display_text = "プライマリディスプレイ"

        self.icon.menu = Menu(
            MenuItem(
                f'インターバル（秒）: {self.controller.config.get("interval", 300)}',
                Menu(
                    MenuItem('+10秒', self.increment_interval(10)),
                    MenuItem('-10秒', self.decrement_interval(10)),
                    MenuItem('30秒', self.set_interval(30)),
                    MenuItem('60秒', self.set_interval(60)),
                    MenuItem('120秒', self.set_interval(120)),
                    MenuItem('300秒', self.set_interval(300)),
                    MenuItem('600秒', self.set_interval(600)),
                )
            ),
            MenuItem('画像/動画を選ぶ', self.choose_file),
            MenuItem(
                f'表示ディスプレイ: {display_text}',
                Menu(*self.get_display_menu_items())
            ),
            MenuItem(mute_text, self.toggle_mute_setting),
            MenuItem(video_suppress_text, self.toggle_video_suppress_setting),
            MenuItem(presentation_enabled_text,
                     self.toggle_presentation_mode_setting),
            # MenuItem(pause_text, self.toggle_pause),
            MenuItem('終了', self.on_quit)
        )

    def choose_file(self, icon, item):
        try:
            # Windows標準のファイルダイアログを使用
            try:
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
