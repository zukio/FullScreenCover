import psutil
import os
import sys
import tempfile
import time

try:
    import win32gui
    import win32con
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# CPU測定結果のキャッシュ
_video_playing_cache = {
    'timestamp': 0,
    'result': False,
    'cache_duration': 1.5  # 1.5秒間キャッシュ
}

if not WIN32_AVAILABLE:
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

        # フルスクリーンかどうかを判定（複数の方法を試す）
        # 1. 基本的な判定（20px誤差許容）
        is_fullscreen_basic = (width >= screen_width -
                               20 and height >= screen_height - 20)

        # 2. Windows 11のメディアプレーヤー特別判定
        is_media_player_fullscreen = False
        if class_name == 'ApplicationFrameWindow' and 'メディア' in title:
            # メディアプレーヤーの場合、より柔軟な判定
            is_media_player_fullscreen = (
                width >= screen_width - 50 and height >= screen_height - 50)

        # 3. ウィンドウの状態を確認
        is_maximized = False
        try:
            window_placement = win32gui.GetWindowPlacement(hwnd)
            show_state = window_placement[1]
            is_maximized = (show_state == 3)  # SW_MAXIMIZE
        except:
            pass

        # 最終的なフルスクリーン判定
        is_fullscreen = is_fullscreen_basic or is_media_player_fullscreen or (
            is_maximized and width >= screen_width - 100)

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
        'vlc', 'mpc-hc', 'mpc-be', 'kmplayer', 'potplayer', 'mpv',
        'windows media player', 'groove', 'movies & tv',
        'メディア プレーヤー', 'media player', 'video player',
        'netflix', 'amazon prime', 'hulu', 'disney+',
        'twitch', 'crunchyroll', 'funimation',
        # その他の人気プレイヤーを追加
        'gom player', 'smplayer', 'bsplayer', 'winamp',
        'kodi', 'plex', 'jellyfin', 'emby'
    ]

    for player in video_players:
        if player in title or player in class_name:
            return True

    # Windows 11のメディアプレーヤーの特別処理
    if (class_name == 'applicationframewindow' and
            ('メディア' in title or 'media' in title or 'player' in title)):
        return True

    return False


def is_youtube_fullscreen(window_info):
    """ブラウザでの動画サービスのフルスクリーンかどうかを判定"""
    if not window_info:
        return False

    title = window_info['title'].lower()
    class_name = window_info['class_name'].lower()

    # ブラウザのクラス名
    browser_classes = ['chrome_widgetwin',
                       'mozillawindowclass', 'applicationframewindow']

    is_browser = any(
        browser_class in class_name for browser_class in browser_classes)

    # 動画サービスのキーワード
    video_services = [
        'youtube', 'netflix', 'amazon prime', 'hulu', 'disney+',
        'twitch', 'crunchyroll', 'funimation', 'abematv', 'nicovideo',
        'bilibili', 'vimeo', 'dailymotion'
    ]

    has_video_service = any(service in title for service in video_services)

    return is_browser and has_video_service and window_info['is_fullscreen']


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


def is_video_playing():
    """動画が実際に再生中かどうかを判定（より確実な検出 + キャッシュ）"""
    global _video_playing_cache

    try:
        current_time = time.time()

        # キャッシュが有効かチェック
        if (current_time - _video_playing_cache['timestamp']) < _video_playing_cache['cache_duration']:
            return _video_playing_cache['result']

        # CPU使用率ベースの判定（より緩い条件で判定）
        target_processes = []

        # まず対象プロセスを特定してCPU測定を開始
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name']:
                    name = proc.info['name'].lower()
                    # 動画プレイヤープロセスの場合
                    if any(player in name for player in ['vlc', 'mpc-hc', 'kmplayer', 'potplayer', 'mpv']):
                        # 初回測定で準備
                        proc.cpu_percent()
                        target_processes.append((proc, name, 3.0))  # 閾値を3%に下げる
                    # Windowsメディアプレーヤー系プロセス（より緩い条件）
                    elif 'media' in name and 'player' in name:
                        # 初回測定で準備
                        proc.cpu_percent()
                        target_processes.append(
                            (proc, name, 1.5))  # 閾値を1.5%に下げる
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, PermissionError):
                continue

        if not target_processes:
            # キャッシュを更新
            _video_playing_cache['timestamp'] = current_time
            _video_playing_cache['result'] = False
            return False

        # より長い待機時間でCPU使用率を測定
        time.sleep(0.5)  # 500ms待機（より長時間測定）

        # 複数回測定して信頼性を向上
        detection_results = []

        for measurement in range(2):  # 2回測定
            cpu_detection_positive = False

            for proc, name, threshold in target_processes:
                try:
                    cpu_percent = proc.cpu_percent()
                    if cpu_percent > threshold:
                        print(f"動画再生検出: {name} (CPU: {cpu_percent:.1f}%)")
                        cpu_detection_positive = True
                        break
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, PermissionError):
                    continue

            detection_results.append(cpu_detection_positive)

            if measurement < 1:  # 最後の測定でない場合は少し待機
                time.sleep(0.2)

        # 2回中1回でも検出されれば再生中と判断
        final_result = any(detection_results)

        # 最終判定結果をキャッシュに保存
        _video_playing_cache['timestamp'] = current_time
        _video_playing_cache['result'] = final_result

        return final_result

    except Exception as e:
        print(f"動画再生状態取得エラー: {e}")
        # エラー時もキャッシュを更新
        _video_playing_cache['timestamp'] = time.time()
        _video_playing_cache['result'] = False
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


def should_suppress_screensaver(suppress_large_window=False):
    """スクリーンセーバーを抑制すべきかどうかを判定

    Args:
        suppress_large_window (bool): 大きなウィンドウでも抑制するかどうか
    """
    try:
        window_info = get_foreground_window_info()
        if not window_info:
            return False

        print(f"フォアグラウンドウィンドウ: {window_info['title']}")
        print(f"クラス名: {window_info['class_name']}")
        print(f"フルスクリーン: {window_info['is_fullscreen']}")

        # 動画プレイヤーウィンドウかどうかの判定
        is_video_player = is_video_player_window(window_info)

        # 動画再生中かどうかを一度だけ確認（CPU測定は一度だけ行う）
        video_playing = is_video_playing()

        # 1. 動画プレイヤーがフルスクリーンかつ動画再生中の場合（正しい抑制条件）
        if is_video_player and window_info['is_fullscreen']:
            if video_playing:
                print("動画プレイヤー（フルスクリーン + 再生中）検出 -> スクリーンセーバー抑制")
                return True
            else:
                print("動画プレイヤー（フルスクリーン）検出されましたが、動画再生中ではないため抑制しません")

        # 2. 動画プレイヤーがフルスクリーンでなくても特定の条件で抑制（設定により有効）
        if suppress_large_window and is_video_player:
            # ウィンドウサイズが大きい場合も抑制（最大化など）
            if window_info['width'] > 1200 and window_info['height'] > 800:
                if video_playing:
                    print("動画プレイヤー（大きなウィンドウ + 再生中）検出 -> スクリーンセーバー抑制")
                    return True
                else:
                    print("動画プレイヤー（大きなウィンドウ）検出されましたが、動画再生中ではないため抑制しません")

        # 3. ブラウザでの動画サービスのフルスクリーン再生の場合
        if is_youtube_fullscreen(window_info):
            print("ブラウザ動画サービス（フルスクリーン）検出 -> スクリーンセーバー抑制")
            return True

        # 4. PowerPointのスライドショーでCPU使用率が高い場合
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
