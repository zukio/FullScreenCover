import pystray
from pystray import MenuItem as item
from PIL import Image
import threading
import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(__file__), '..', 'config.json')


class TrayApp:
    def __init__(self, config):
        self.config = config
        self.icon = None
        self.running = True

    def setup_tray(self):
        image = Image.open(os.path.join(os.path.dirname(
            __file__), '..', 'assets', 'icon.ico'))
        menu = (
            item('設定変更', self.on_config),
            item('終了', self.on_exit)
        )
        self.icon = pystray.Icon('TrayApp', image, 'TrayApp', menu)

    def on_config(self, icon, item):
        # 設定変更ダイアログ（簡易版）
        # ここで設定変更処理を実装
        self.config['example'] = not self.config.get('example', False)
        self.save_config()

    def save_config(self):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)

    def on_exit(self, icon, item):
        self.running = False
        self.icon.stop()

    def run(self):
        self.setup_tray()
        self.icon.run()

# トレイアプリ起動用関数


def start_tray_app(config):
    app = TrayApp(config)
    tray_thread = threading.Thread(target=app.run)
    tray_thread.start()
    return app
