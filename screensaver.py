# フルスクリーン表示処理（画像/動画切替）
import tkinter as tk
import cv2
import os
import sys
from threading import Thread
from PIL import Image, ImageTk
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


def show_screensaver(media_file):
    # media_fileが相対パスの場合、リソースパスとして解決
    if not os.path.isabs(media_file):
        media_file = get_resource_path(media_file)

    ext = os.path.splitext(media_file)[1].lower()
    if ext in ['.png', '.jpg', '.jpeg', '.bmp']:
        show_image(media_file)
    elif ext in ['.mp4', '.avi', '.mov']:
        show_video(media_file)


def show_image(image_path):
    # 毎回新しいtkinterインスタンスを作成
    root = tk.Tk()
    root.attributes('-fullscreen', True)
    root.config(cursor='none')
    root.configure(bg='black')

    # クリーンアップフラグ
    closed = {'flag': False}

    try:
        # PILを使って画像を読み込み、tkinter用に変換
        pil_image = Image.open(image_path)

        # 画面サイズを取得
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

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
        # 確実にクリーンアップ
        try:
            if root.winfo_exists():
                root.destroy()
        except:
            pass


def show_video(video_path):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"Video file could not be opened: {video_path}")
        return

    cv2.namedWindow('screensaver', cv2.WND_PROP_FULLSCREEN)
    cv2.setWindowProperty(
        'screensaver', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

    # グローバルフック用のフラグ
    closed = {'flag': False}

    def global_close(*args):
        if not closed['flag']:
            closed['flag'] = True
            print("グローバルフック発火: ビデオ再生中断 detected")
            cap.release()
            cv2.destroyAllWindows()

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
            cv2.imshow('screensaver', frame)
            # ESCキーまたは任意のキーで終了
            if cv2.waitKey(30) & 0xFF != 255:  # 何かキーが押された
                break
    except Exception as e:
        print(f"Video playback error: {e}")
    finally:
        cap.release()
        cv2.destroyAllWindows()
