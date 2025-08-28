import threading

try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
    from pycaw.constants import AudioSessionState
    from comtypes import CLSCTX_ALL
    import comtypes
    import ctypes
    PYCAW_AVAILABLE = True
except ImportError:
    PYCAW_AVAILABLE = False


def list_input_devices():
    """Return a list of tuples (index, name) for available input devices."""
    if not PYAUDIO_AVAILABLE:
        print("pyaudio not available")
        return []

    pa = pyaudio.PyAudio()
    devices = []
    try:
        for i in range(pa.get_device_count()):
            info = pa.get_device_info_by_index(i)
            if info.get("maxInputChannels", 0) > 0:
                devices.append((i, info.get("name")))
    finally:
        pa.terminate()
    return devices


def get_device_name(index: int) -> str | None:
    """Return the device name for the given index or None if not found."""
    if not PYAUDIO_AVAILABLE:
        return None

    pa = pyaudio.PyAudio()
    try:
        info = pa.get_device_info_by_index(index)
        return info.get("name")
    except Exception:
        return None
    finally:
        pa.terminate()


class VolumeController:
    """システム音量制御クラス"""

    def __init__(self):
        self.volume_interface = None
        self.previous_mute_state = None
        self.previous_volume = None
        self._com_initialized = False
        self._initialize()

    def _initialize(self):
        """音量制御インターフェースを初期化"""
        if not PYCAW_AVAILABLE:
            print("pycaw not available. Volume control disabled.")
            return

        try:
            # 追加: COM初期化（多重初期化を避ける）
            try:
                comtypes.CoInitialize()
                self._com_initialized = True
            except Exception:
                # 既に初期化済みなどは無視
                pass
            # デフォルトの音声出力デバイスを取得
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume_interface = ctypes.cast(
                interface, ctypes.POINTER(IAudioEndpointVolume))
        except Exception as e:
            print(f"音量制御初期化エラー: {e}")
            self.volume_interface = None

    def __del__(self):
        # 追加: COM後始末
        try:
            if self._com_initialized:
                comtypes.CoUninitialize()
        except Exception:
            pass

    # 追加: アクティブな全セッションを取得
    def _get_active_sessions(self):
        try:
            # COMの初期化状態を確認し、必要に応じて初期化
            try:
                comtypes.CoInitialize()
            except Exception:
                # 既に初期化済みの場合は無視
                pass

            sessions = AudioUtilities.GetAllSessions()
            # AudioSessionState.Activeを使用（値は1）
            return [s for s in sessions if s.State == AudioSessionState.Active and s.SimpleAudioVolume is not None]
        except Exception as e:
            print(f"セッション取得エラー: {e}")
            return []

    def get_mute_state(self):
        """現在のミュート状態を取得"""
        if not self.volume_interface:
            return False
        try:
            return self.volume_interface.GetMute()
        except Exception as e:
            print(f"ミュート状態取得エラー: {e}")
            return False

    def get_volume(self):
        """現在の音量を取得（0.0-1.0）"""
        if not self.volume_interface:
            return 0.5
        try:
            return self.volume_interface.GetMasterVolumeLevelScalar()
        except Exception as e:
            print(f"音量取得エラー: {e}")
            return 0.5

    def set_mute(self, mute=True):
        """ミュート状態を設定"""
        if not self.volume_interface:
            print("音量制御が利用できません")
            return False
        try:
            self.volume_interface.SetMute(mute, None)
            return True
        except Exception as e:
            print(f"ミュート設定エラー: {e}")
            return False

    def set_volume(self, volume):
        """音量を設定（0.0-1.0）"""
        if not self.volume_interface:
            print("音量制御が利用できません")
            return False
        try:
            self.volume_interface.SetMasterVolumeLevelScalar(volume, None)
            return True
        except Exception as e:
            print(f"音量設定エラー: {e}")
            return False

    def save_current_state(self):
        """現在の音量・ミュート状態を保存（システムレベルのみ）"""
        # 既に保存済みの場合は再保存しない
        if self.previous_mute_state is not None:
            print("音量状態は既に保存済みです")
            return

        # システムレベルの状態のみ保存
        self.previous_mute_state = self.get_mute_state()
        self.previous_volume = self.get_volume()

        print(
            f"音量状態保存: 音量={self.previous_volume:.2f}, ミュート={self.previous_mute_state}")

    def restore_previous_state(self):
        """保存された音量・ミュート状態を復元（システムレベルのみ）"""
        if self.previous_mute_state is not None:
            self.set_mute(self.previous_mute_state)
        if self.previous_volume is not None:
            self.set_volume(self.previous_volume)
        print(
            f"音量状態復元: 音量={self.previous_volume:.2f}, ミュート={self.previous_mute_state}")

        # 復元後は状態をクリア
        self.previous_mute_state = None
        self.previous_volume = None

    def mute_for_screensaver(self):
        """スクリーンセーバー用にミュート（システムレベルのミュートのみ）"""
        # 現在の状態を保存
        self.save_current_state()

        # システムレベルのミュートを実行
        # これにより、既存および新規のすべてのアプリケーションがミュートされる
        success = self.set_mute(True)

        if success and self.get_mute_state():
            print("スクリーンセーバー: システムレベルミュート有効")
        else:
            # フォールバック: システム音量を0にする
            volume_success = self.set_volume(0.0)
            if volume_success:
                print("スクリーンセーバー: システム音量を0に設定（フォールバック）")
            else:
                print("スクリーンセーバー: ミュート失敗")
                return False

        return True

    def unmute_after_screensaver(self):
        """スクリーンセーバー終了後にミュート解除（元の状態に復元）"""
        self.restore_previous_state()
        print("スクリーンセーバー終了: 音声を復元しました")
