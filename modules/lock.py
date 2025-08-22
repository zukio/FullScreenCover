import psutil
import os
import sys
import tempfile

try:
    import win32gui
    import win32con
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("win32 modules not available. Some features may be disabled.")


class SingleInstance:
    def __init__(self, lockname):
        self.lockfile = os.path.join(tempfile.gettempdir(), f'{lockname}.lock')
        self.fp = None

    def acquire(self):
        if os.path.exists(self.lockfile):
            print('多重起動防止: 既に起動しています。')
            sys.exit(1)
        self.fp = open(self.lockfile, 'w')
        self.fp.write(str(os.getpid()))
        self.fp.flush()

    def release(self):
        if self.fp:
            self.fp.close()
        if os.path.exists(self.lockfile):
            os.remove(self.lockfile)


def get_foreground_window_info():
    """フォアグラウンドウィンドウの情報を取得"""
    if not WIN32_AVAILABLE:
        print("win32 modules not available")
        return None

    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)

        # ウィンドウの位置とサイズを取得
        rect = win32gui.GetWindowRect(hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]

        # スクリーンサイズを取得
        screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

        # フルスクリーンかどうかを判定
        is_fullscreen = (width >= screen_width -
                         20 and height >= screen_height - 20)

        return {
            'hwnd': hwnd,
            'title': title,
            'class_name': class_name,
            'width': width,
            'height': height,
            'is_fullscreen': is_fullscreen,
            'rect': rect
        }
    except Exception as e:
        print(f"ウィンドウ情報取得エラー: {e}")
        return None


def is_video_player_window(window_info):
    """動画プレイヤーウィンドウかどうかを判定"""
    if not window_info:
        return False

    title = window_info['title'].lower()
    class_name = window_info['class_name'].lower()

    # 明確な動画プレイヤー
    video_players = [
        'vlc', 'mpc-hc', 'mpc-be', 'kmplayer', 'potplayer',
        'windows media player', 'groove', 'movies & tv'
    ]

    for player in video_players:
        if player in title or player in class_name:
            return True

    return False


def is_youtube_fullscreen(window_info):
    """YouTubeのフルスクリーンかどうかを判定"""
    if not window_info:
        return False

    title = window_info['title'].lower()
    class_name = window_info['class_name'].lower()

    # ブラウザのクラス名
    browser_classes = ['chrome_widgetwin',
                       'mozillawindowclass', 'applicationframewindow']

    is_browser = any(
        browser_class in class_name for browser_class in browser_classes)
    has_youtube = 'youtube' in title

    return is_browser and has_youtube and window_info['is_fullscreen']


def is_powerpoint_slideshow_running():
    """PowerPointのスライドショーが実行中かどうかを判定"""
    try:
        window_info = get_foreground_window_info()
        if not window_info:
            return False

        title = window_info['title'].lower()
        class_name = window_info['class_name'].lower()

        # PowerPointのスライドショーウィンドウ
        powerpoint_slideshow_classes = ['pptframeclass', 'screenclass']

        is_powerpoint = ('powerpoint' in title or
                         any(ppt_class in class_name for ppt_class in powerpoint_slideshow_classes))

        return is_powerpoint and window_info['is_fullscreen']
    except Exception as e:
        print(f"PowerPoint判定エラー: {e}")
        return False


def get_cpu_usage_for_process(process_name):
    """指定プロセスのCPU使用率を取得"""
    try:
        for proc in psutil.process_iter(['name', 'cpu_percent']):
            if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                # 短時間でCPU使用率を測定
                cpu_percent = proc.cpu_percent(interval=0.5)
                return cpu_percent
        return 0
    except Exception as e:
        print(f"CPU使用率取得エラー: {e}")
        return 0


def is_powerpoint_high_cpu():
    """PowerPointのCPU使用率が高いかどうかを判定"""
    cpu_usage = get_cpu_usage_for_process('powerpnt')
    print(f"PowerPoint CPU使用率: {cpu_usage:.1f}%")
    return cpu_usage > 10  # 10%以上をしきい値とする


def should_suppress_screensaver():
    """スクリーンセーバーを抑制すべきかどうかを判定"""
    try:
        window_info = get_foreground_window_info()
        if not window_info:
            return False

        print(f"フォアグラウンドウィンドウ: {window_info['title']}")
        print(f"クラス名: {window_info['class_name']}")
        print(f"フルスクリーン: {window_info['is_fullscreen']}")

        # 1. 動画プレイヤーがフルスクリーンの場合
        if is_video_player_window(window_info) and window_info['is_fullscreen']:
            print("動画プレイヤー（フルスクリーン）検出 -> スクリーンセーバー抑制")
            return True

        # 2. YouTubeのフルスクリーン再生の場合
        if is_youtube_fullscreen(window_info):
            print("YouTube（フルスクリーン）検出 -> スクリーンセーバー抑制")
            return True

        # 3. PowerPointのスライドショーでCPU使用率が高い場合
        if is_powerpoint_slideshow_running():
            if is_powerpoint_high_cpu():
                print("PowerPoint（スライドショー + 高CPU）検出 -> スクリーンセーバー抑制")
                return True
            else:
                print("PowerPoint（スライドショー）検出されましたが、CPU使用率が低いため抑制しません")

        return False

    except Exception as e:
        print(f"スクリーンセーバー抑制判定エラー: {e}")
        return False
