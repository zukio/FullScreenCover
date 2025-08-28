# フルスクリーン表示処理（画像/動画切替）
import tkinter as tk
import cv2
import os
import sys
import threading
from threading import Thread
from PIL import Image, ImageTk
from modules.utils.display_utils import get_display_by_index, get_primary_display, DisplayInfo
from modules.utils.cursor_control import hide_system_cursor, show_system_cursor
from typing import Optional, List
try:
    from pynput import mouse, keyboard
except ImportError:
    mouse = None
    keyboard = None


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def show_screensaver(media_file, display_index: Optional[int] = None):
    """
    指定されたディスプレイでスクリーンセーバーを表示

    Args:
        media_file: 表示するメディアファイルのパス
        display_index: 表示するディスプレイのインデックス（Noneの場合はプライマリディスプレイ）
    """
    # media_fileが相対パスの場合、リソースパスとして解決
    if not os.path.isabs(media_file):
        media_file = get_resource_path(media_file)

    # ディスプレイ情報を取得
    if display_index is not None:
        display_info = get_display_by_index(display_index)
        if display_info is None:
            print(
                f"指定されたディスプレイ（インデックス: {display_index}）が見つかりません。プライマリディスプレイを使用します。")
            display_info = get_primary_display()
    else:
        display_info = get_primary_display()

    print(f"ディスプレイ {display_info.index + 1} でスクリーンセーバーを表示: {display_info}")

    ext = os.path.splitext(media_file)[1].lower()
    if ext in ['.png', '.jpg', '.jpeg', '.bmp']:
        show_image(media_file, display_info)
    elif ext in ['.mp4', '.avi', '.mov']:
        show_video(media_file, display_info)


def show_image(image_path, display_info: DisplayInfo):
    """
    指定されたディスプレイで画像を表示

    Args:
        image_path: 画像ファイルのパス
        display_info: 表示するディスプレイの情報
    """
    print(f"ディスプレイ {display_info.index + 1} で画像表示開始: {display_info.width}x{display_info.height} at ({display_info.x}, {display_info.y})")

    # 毎回新しいtkinterインスタンスを作成
    root = tk.Tk()

    # システムレベルでマウスカーソーを非表示（tkinterのcursor='none'に加えて）
    cursor_hidden = hide_system_cursor()
    print(f"画像表示開始: マウスカーソー非表示 = {cursor_hidden}")

    # 基本設定
    root.config(cursor='none')
    root.configure(bg='black')

    # ウィンドウの装飾を除去（フルスクリーンの代替）
    root.overrideredirect(True)

    # 指定されたディスプレイに正確に配置
    geometry_string = f"{display_info.width}x{display_info.height}+{display_info.x}+{display_info.y}"
    print(f"ジオメトリ設定: {geometry_string}")
    root.geometry(geometry_string)

    # 最前面表示
    root.attributes('-topmost', True)
    root.focus_force()

    # ウィンドウサイズと位置の確認
    root.update()
    print(f"設定後ウィンドウ位置: {root.winfo_x()}, {root.winfo_y()}")
    print(f"設定後ウィンドウサイズ: {root.winfo_width()}x{root.winfo_height()}")

    # クリーンアップフラグ
    closed = {'flag': False}

    try:
        # PILを使って画像を読み込み、tkinter用に変換
        pil_image = Image.open(image_path)

        # 指定されたディスプレイのサイズを使用
        screen_width = display_info.width
        screen_height = display_info.height

        # 画像をリサイズ（アスペクト比を保持）
        pil_image.thumbnail((screen_width, screen_height),
                            Image.Resampling.LANCZOS)

        # tkinter用のImageTkオブジェクトに変換
        img = ImageTk.PhotoImage(pil_image)

        label = tk.Label(root, image=img, bg='black')
        label.pack(expand=True)

        # 画像の参照を保持
        label.image = img
    except (tk.TclError, FileNotFoundError, Exception) as e:
        print(f"Image loading error: {e}")
        root.destroy()
        return

    def close(event=None):
        if not closed['flag']:
            closed['flag'] = True
            try:
                # マウスカーソーを復元
                if cursor_hidden:
                    show_system_cursor()
                    print("画像表示終了: マウスカーソー復元")

                # グローバルフックリスナーを停止
                if hasattr(close, 'mouse_listener'):
                    try:
                        close.mouse_listener.stop()
                    except:
                        pass
                if hasattr(close, 'keyboard_listener'):
                    try:
                        close.keyboard_listener.stop()
                    except:
                        pass

                root.quit()  # mainloopを終了
                root.destroy()
            except Exception as e:
                print(f"close error: {e}")

    # マウス・キー操作で即解除
    root.bind_all('<Motion>', close)
    root.bind_all('<Key>', close)
    root.bind_all('<Button>', close)

    # グローバルフック（ウィンドウ外の操作も検知）
    def global_close(*args):
        if not closed['flag']:
            print("グローバルフック発火: マウス・キーボード操作 detected")
            close()

    def mouse_hook():
        if mouse:
            try:
                listener = mouse.Listener(
                    on_move=global_close, on_click=global_close, on_scroll=global_close)
                close.mouse_listener = listener  # リスナーを保存
                with listener:
                    listener.join()
            except Exception as e:
                print(f"mouse_hook error: {e}")

    def keyboard_hook():
        if keyboard:
            try:
                listener = keyboard.Listener(on_press=global_close)
                close.keyboard_listener = listener  # リスナーを保存
                with listener:
                    listener.join()
            except Exception as e:
                print(f"keyboard_hook error: {e}")

    if mouse:
        Thread(target=mouse_hook, daemon=True).start()
    if keyboard:
        Thread(target=keyboard_hook, daemon=True).start()

    # 自動終了タイマーを削除（ユーザー操作のみで終了するように変更）
    # root.after(10000, close)  # この自動終了が問題の原因

    try:
        root.mainloop()
    except Exception as e:
        print(f"mainloop error: {e}")
    finally:
        # マウスカーソーを復元
        if cursor_hidden:
            show_system_cursor()
            print("画像表示終了: マウスカーソー復元 (finally)")

        # 確実にクリーンアップ
        try:
            if root.winfo_exists():
                root.destroy()
        except:
            pass


def show_video(video_path, display_info: DisplayInfo):
    """
    指定されたディスプレイで動画を表示

    Args:
        video_path: 動画ファイルのパス
        display_info: 表示するディスプレイの情報
    """
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"Video file could not be opened: {video_path}")
        return

    # システムレベルでマウスカーソーを非表示
    cursor_hidden = hide_system_cursor()
    print(f"動画表示開始: マウスカーソー非表示 = {cursor_hidden}")

    # ウィンドウ名にディスプレイ情報を含める
    window_name = f'screensaver_display_{display_info.index}'
    cv2.namedWindow(window_name, cv2.WND_PROP_FULLSCREEN)
    cv2.setWindowProperty(
        window_name, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

    # 指定されたディスプレイに配置
    cv2.moveWindow(window_name, display_info.x, display_info.y)

    # 最前面表示を設定
    cv2.setWindowProperty(window_name, cv2.WND_PROP_TOPMOST, 1)

    # グローバルフック用のフラグ
    closed = {'flag': False}

    def global_close(*args):
        if not closed['flag']:
            closed['flag'] = True
            print("グローバルフック発火: ビデオ再生中断 detected")
            # マウスカーソーを復元
            if cursor_hidden:
                show_system_cursor()
                print("動画表示終了: マウスカーソー復元")
            cap.release()
            cv2.destroyWindow(window_name)

    def mouse_hook():
        if mouse:
            try:
                listener = mouse.Listener(
                    on_move=global_close, on_click=global_close, on_scroll=global_close)
                with listener:
                    listener.join()
            except Exception as e:
                print(f"video mouse_hook error: {e}")

    def keyboard_hook():
        if keyboard:
            try:
                listener = keyboard.Listener(on_press=global_close)
                with listener:
                    listener.join()
            except Exception as e:
                print(f"video keyboard_hook error: {e}")

    if mouse:
        Thread(target=mouse_hook, daemon=True).start()
    if keyboard:
        Thread(target=keyboard_hook, daemon=True).start()

    try:
        while cap.isOpened() and not closed['flag']:
            ret, frame = cap.read()
            if not ret:
                break
            cv2.imshow(window_name, frame)
            # ESCキーまたは任意のキーで終了
            if cv2.waitKey(30) & 0xFF != 255:  # 何かキーが押された
                break
    except Exception as e:
        print(f"Video playback error: {e}")
    finally:
        # マウスカーソーを復元
        if cursor_hidden:
            show_system_cursor()
            print("動画表示終了: マウスカーソー復元 (finally)")
        cap.release()
        cv2.destroyWindow(window_name)


def show_screensaver_on_all_displays_simultaneously(media_file, display_list: List[DisplayInfo]):
    """
    複数ディスプレイで同時にスクリーンセーバーを表示

    Args:
        media_file: 表示するメディアファイル
        display_list: 表示対象のディスプレイリスト
    """
    if not display_list:
        print("表示対象のディスプレイがありません")
        return

    print(f"同時表示開始: {len(display_list)}台のディスプレイ")

    # 各ディスプレイ用のプロセスを作成（tkinterの制限回避）
    import subprocess
    import sys

    processes = []

    try:
        for display_info in display_list:
            # 各ディスプレイ用に独立したPythonプロセスを起動
            cmd = [
                sys.executable, '-c',
                f'''
import sys
import os
sys.path.append("{os.getcwd().replace(chr(92), chr(92)+chr(92))}")
from screensaver import show_image
from modules.utils.display_utils import DisplayInfo

display = DisplayInfo({display_info.index}, {display_info.x}, {display_info.y}, {display_info.width}, {display_info.height}, {display_info.is_primary})
show_image(r"{media_file}", display)
'''
            ]

            print(f"ディスプレイ {display_info.index + 1} 用プロセス起動中...")
            process = subprocess.Popen(cmd, cwd=os.getcwd())
            processes.append(process)
            print(
                f"ディスプレイ {display_info.index + 1} 用プロセス起動完了 (PID: {process.pid})")

        print("全ディスプレイでの表示開始完了、終了待機中...")

        # 全プロセスの完了を待機
        for i, process in enumerate(processes):
            process.wait()
            print(f"ディスプレイ {i + 1} プロセス終了")

    except Exception as e:
        print(f"同時表示エラー: {e}")
        # エラー時は全プロセスを強制終了
        for process in processes:
            try:
                process.terminate()
            except:
                pass

    print("全ディスプレイでの同時表示完了")
