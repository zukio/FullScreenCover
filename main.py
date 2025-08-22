import os
import sys
import threading
import json
import time
from modules.utils.path_utils import get_idle_duration
from modules.lock import should_suppress_screensaver
from modules.audio_devices import VolumeController
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
        print("TrayMenu初期化開始")
        self.tray = TrayMenu(self)
        self.running = True
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
                'mute_on_screensaver': True
            }
            self.save_config()

        # デフォルト値の設定
        if 'mute_on_screensaver' not in self.config:
            self.config['mute_on_screensaver'] = True
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
                suppress = should_suppress_screensaver()
                mute_setting = self.config.get('mute_on_screensaver', False)
                print(
                    f"idle={idle:.1f}, interval={self.config['interval']}, suppress={suppress}, showing={self.showing}, mute={mute_setting}")

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
        print("停止処理開始")
        self.running = False
        if hasattr(self, 'tray') and self.tray:
            self.tray.stop()

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
