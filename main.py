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


CONFIG_PATH = get_resource_path('config.json')


class ScreensaverController:
    def __init__(self):
        print("ScreensaverController初期化開始")
        self.load_config()
        print("VolumeController初期化開始")
        self.volume_controller = VolumeController()
        print("PresentationModeController初期化開始")
        # サイレントモード設定を取得
        silent_mode = self.config.get('presentation_mode_silent', False)
        # 高度な機能設定を取得
        features_config = self.config.get('presentation_features', {})
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
                'interval': 5,
                'media_file': get_resource_path('assets/image.png'),
                'mute_on_screensaver': True,
                'suppress_during_video': True,
                'enable_presentation_mode': False,
                'presentation_mode_silent': True,
                'presentation_features': {
                    'disable_screensaver': True,
                    'prevent_sleep': True,
                    'block_notifications': True,
                    'replace_wallpaper': False
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

        if 'enable_presentation_mode' not in self.config:
            self.config['enable_presentation_mode'] = False
            self.save_config()

        if 'presentation_mode_silent' not in self.config:
            self.config['presentation_mode_silent'] = True
            self.save_config()

        if 'presentation_features' not in self.config:
            self.config['presentation_features'] = {
                'disable_screensaver': True,
                'prevent_sleep': True,
                'block_notifications': True,
                'replace_wallpaper': False
            }
            self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"設定保存エラー: {e}")

    def show_screensaver_with_mute(self, media_file):
        """ミュート機能付きスクリーンセーバー表示"""
        muted = False
        presentation_enabled = False
        try:
            # プレゼンテーションモード設定が有効な場合、プレゼンテーションモードを有効にする
            if self.config.get('enable_presentation_mode', False):
                presentation_enabled = self.presentation_controller.enable_presentation_mode()
                if presentation_enabled:
                    print("プレゼンテーションモードを有効にしました")

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

            # プレゼンテーションモードを無効にする
            if presentation_enabled and self.config.get('enable_presentation_mode', False):
                self.presentation_controller.disable_presentation_mode()
                print("プレゼンテーションモードを無効にしました")

    def monitor(self):
        print("モニタリング開始")
        while self.running:
            try:
                idle = get_idle_duration()

                # 動画抑制設定が有効な場合のみチェック
                suppress = False
                if self.config.get('suppress_during_video', True):
                    suppress = should_suppress_screensaver()

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
