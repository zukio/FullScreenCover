import os
import sys
import threading
import json
import time
from modules.utils.path_utils import get_idle_duration
from modules.lock import should_suppress_screensaver
from modules.audio_devices import VolumeController
from modules.presentation_mode import get_presentation_controller
from tray_menu import TrayMenu
from screensaver import show_screensaver


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_config_path():
    """Get config.json path from executable directory, not from _MEIPASS"""
    try:
        # PyInstallerの場合、実行ファイルのディレクトリを取得
        if hasattr(sys, '_MEIPASS'):
            # 実行ファイルのディレクトリを取得
            exe_dir = os.path.dirname(sys.executable)
        else:
            # 開発環境では現在のディレクトリ
            exe_dir = os.path.abspath(".")
        return os.path.join(exe_dir, 'config.json')
    except Exception:
        return os.path.join(os.path.abspath("."), 'config.json')


CONFIG_PATH = get_config_path()


class ScreensaverController:
    def __init__(self):
        print("ScreensaverController初期化開始")
        self.load_config()
        print("VolumeController初期化開始")
        self.volume_controller = VolumeController()
        print("PresentationModeController初期化開始")
        # プレゼンテーションモードの制御
        presentation_enabled = self.config.get(
            'enable_presentation_mode', False)
        if presentation_enabled:
            # プレゼンテーションモードが有効な場合のみ、個別設定を適用
            features_config = self.config.get('presentation_features', {})
            silent_mode = features_config.get('silent_notifications', False)
        else:
            # プレゼンテーションモードが無効な場合、すべての機能を無効化
            features_config = {
                'disable_screensaver': False,
                'prevent_sleep': False,
                'block_notifications': False
            }
            silent_mode = False

        self.presentation_controller = get_presentation_controller(
            silent_mode=silent_mode, features_config=features_config)
        print("TrayMenu初期化開始")
        self.tray = TrayMenu(self)
        self.running = True
        self.stopping = False  # 停止処理フラグを追加
        self.showing = False
        self.monitor_thread = threading.Thread(
            target=self.monitor, daemon=True)
        self.monitor_thread.start()
        print("ScreensaverController初期化完了")

    def load_config(self):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            # デフォルト設定
            self.config = {
                'interval': 60,
                'media_file': get_resource_path('assets/image.png'),
                'mute_on_screensaver': True,
                'suppress_during_video': True,
                'suppress_large_window': False,  # 大きなウィンドウでの抑制（デフォルトOFF）
                'enable_presentation_mode': False,
                'presentation_features': {
                    'disable_screensaver': True,
                    'prevent_sleep': True,
                    'silent_notifications': True,
                    'block_notifications': True
                }
            }
            self.save_config()

        # デフォルト値の設定
        if 'mute_on_screensaver' not in self.config:
            self.config['mute_on_screensaver'] = True
            self.save_config()

        if 'suppress_during_video' not in self.config:
            self.config['suppress_during_video'] = True
            self.save_config()

        if 'suppress_large_window' not in self.config:
            self.config['suppress_large_window'] = False
            self.save_config()

        if 'enable_presentation_mode' not in self.config:
            self.config['enable_presentation_mode'] = False
            self.save_config()

        # 古い設定項目からのマイグレーション
        if 'presentation_mode_noticesilent' in self.config:
            # 古い設定を新しい構造に移行
            old_value = self.config.pop('presentation_mode_noticesilent')
            if 'presentation_features' not in self.config:
                self.config['presentation_features'] = {}
            self.config['presentation_features']['silent_notifications'] = old_value
            self.save_config()

        if 'presentation_features' not in self.config:
            self.config['presentation_features'] = {
                'disable_screensaver': True,
                'prevent_sleep': True,
                'silent_notifications': True,
                'block_notifications': True
            }
            self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            # 設定保存後、プレゼンテーションコントローラーを再初期化
            self.reinitialize_presentation_controller()
        except Exception as e:
            print(f"設定保存エラー: {e}")

    def reinitialize_presentation_controller(self):
        """プレゼンテーションコントローラーの再初期化"""
        try:
            # プレゼンテーションモードの制御
            presentation_enabled = self.config.get(
                'enable_presentation_mode', False)
            if presentation_enabled:
                # プレゼンテーションモードが有効な場合のみ、個別設定を適用
                features_config = self.config.get('presentation_features', {})
                silent_mode = features_config.get(
                    'silent_notifications', False)
            else:
                # プレゼンテーションモードが無効な場合、すべての機能を無効化
                features_config = {
                    'disable_screensaver': False,
                    'prevent_sleep': False,
                    'block_notifications': False,
                }
                silent_mode = False

            # 新しい設定でコントローラーを再作成
            self.presentation_controller = get_presentation_controller(
                silent_mode=silent_mode, features_config=features_config)
            print(
                f"プレゼンテーションコントローラーを再初期化しました (enabled: {presentation_enabled}, silent: {silent_mode})")
        except Exception as e:
            print(f"プレゼンテーションコントローラー再初期化エラー: {e}")

    def show_screensaver_with_mute(self, media_file):
        """ミュート機能付きスクリーンセーバー表示"""
        muted = False
        try:
            # ミュート設定が有効な場合、音声をミュート
            if self.config.get('mute_on_screensaver', False):
                muted = self.volume_controller.mute_for_screensaver()

            # 画像ファイルのパスを解決
            resolved_media_file = media_file
            if not os.path.isabs(media_file):
                resolved_media_file = get_resource_path(media_file)
            elif not os.path.exists(media_file):
                # 絶対パスだが存在しない場合、リソースパスとして再試行
                basename = os.path.basename(media_file)
                resolved_media_file = get_resource_path(f'assets/{basename}')

            print(f"スクリーンセーバー表示: {resolved_media_file}")
            # スクリーンセーバーを表示
            show_screensaver(resolved_media_file)

        finally:
            # スクリーンセーバー終了後、音声を復元
            if muted and self.config.get('mute_on_screensaver', False):
                self.volume_controller.unmute_after_screensaver()

    def monitor(self):
        print("モニタリング開始")
        while self.running:
            try:
                idle = get_idle_duration()

                # 動画抑制設定が有効な場合のみチェック
                suppress = False
                if self.config.get('suppress_during_video', True):
                    suppress_large_window = self.config.get(
                        'suppress_large_window', False)
                    suppress = should_suppress_screensaver(
                        suppress_large_window)

                mute_setting = self.config.get('mute_on_screensaver', False)
                video_suppress_setting = self.config.get(
                    'suppress_during_video', True)
                presentation_setting = self.config.get(
                    'enable_presentation_mode', False)
                print(
                    f"idle={idle:.1f}, interval={self.config['interval']}, suppress={suppress}, showing={self.showing}, mute={mute_setting}, video_suppress={video_suppress_setting}, presentation={presentation_setting}")

                if idle > self.config['interval'] and not suppress and not self.showing:
                    print("スクリーンセーバー表示開始")
                    self.showing = True
                    try:
                        self.show_screensaver_with_mute(
                            self.config['media_file'])
                    except Exception as e:
                        print(f"スクリーンセーバー表示エラー: {e}")
                    finally:
                        self.showing = False
                        print("スクリーンセーバー表示終了")

                time.sleep(2)
            except Exception as e:
                print(f"モニタリングエラー: {e}")
                time.sleep(5)
        print("モニタリング終了")

    def stop(self):
        if self.stopping:
            return  # 既に停止処理中の場合は何もしない

        print("停止処理開始")
        self.stopping = True
        self.running = False

        try:
            # プレゼンテーションモードを無効にする
            if hasattr(self, 'presentation_controller') and self.presentation_controller:
                if self.presentation_controller.is_presentation_mode_active():
                    self.presentation_controller.disable_presentation_mode()
                    print("プレゼンテーションモードを無効にしました")
        except Exception as e:
            print(f"プレゼンテーションモード停止エラー: {e}")

        try:
            if hasattr(self, 'tray') and self.tray:
                self.tray.stop()
        except Exception as e:
            print(f"トレイ停止エラー: {e}")

        print("停止処理完了")

    def run(self):
        try:
            if hasattr(self, 'tray') and self.tray:
                self.tray.run()
        except KeyboardInterrupt:
            print("キーボード割り込み")
        except Exception as e:
            print(f"実行エラー: {e}")
        finally:
            self.stop()


def main():
    try:
        controller = ScreensaverController()
        controller.run()
    except Exception as e:
        print(f"メインエラー: {e}")
    finally:
        print("アプリケーション終了")


if __name__ == '__main__':
    main()
