import threading

try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
    from pycaw.constants import AudioSessionStateActive
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
        # 追加: セッション復元用
        self._touched_sessions = []  # [(session, prev_mute, prev_vol)]
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
            sessions = AudioUtilities.GetAllSessions()
            # AudioSessionStateActiveの代わりに定数値1を使用
            return [s for s in sessions if s.State == 1 and s.SimpleAudioVolume is not None]
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
        """現在の音量・ミュート状態を保存"""
        # 既に保存済みの場合は再保存しない
        if self.previous_mute_state is not None:
            print("音量状態は既に保存済みです")
            return

        self.previous_mute_state = self.get_mute_state()
        self.previous_volume = self.get_volume()
        # 追加: セッション状態も保存
        self._touched_sessions = []
        for s in self._get_active_sessions():
            try:
                prev_m = s.SimpleAudioVolume.GetMute()
                prev_v = s.SimpleAudioVolume.GetMasterVolume()
                self._touched_sessions.append((s, prev_m, prev_v))
            except Exception:
                pass
        print(
            f"音量状態保存: 音量={self.previous_volume:.2f}, ミュート={self.previous_mute_state}")

    def restore_previous_state(self):
        """保存された音量・ミュート状態を復元"""
        if self.previous_mute_state is not None:
            self.set_mute(self.previous_mute_state)
        if self.previous_volume is not None:
            self.set_volume(self.previous_volume)
        print(
            f"音量状態復元: 音量={self.previous_volume:.2f}, ミュート={self.previous_mute_state}")
        # 追加: セッションの復元
        for s, prev_m, prev_v in self._touched_sessions:
            try:
                s.SimpleAudioVolume.SetMute(prev_m, None)
                s.SimpleAudioVolume.SetMasterVolume(prev_v, None)
            except Exception:
                pass
        self._touched_sessions = []

        # 復元後は状態をクリア
        self.previous_mute_state = None
        self.previous_volume = None

    def mute_for_screensaver(self):
        """スクリーンセーバー用にミュート（状態を保存してからミュート）"""
        self.save_current_state()
        ok = self.set_mute(True)
        if not ok or not self.get_mute_state():
            # フォールバック1: エンドポイント音量を0へ
            v_ok = self.set_volume(0.0)
            if not v_ok:
                print("エンドポイント音量0化に失敗")
        # フォールバック2: アクティブ・セッションもミュート/0化
        for s in self._get_active_sessions():
            try:
                s.SimpleAudioVolume.SetMute(True, None)
                s.SimpleAudioVolume.SetMasterVolume(0.0, None)
            except Exception:
                pass
        print("スクリーンセーバー: 音声をミュート（多段フォールバック適用）")
        return True

    def unmute_after_screensaver(self):
        """スクリーンセーバー終了後にミュート解除（元の状態に復元）"""
        self.restore_previous_state()
        print("スクリーンセーバー終了: 音声を復元しました")
