"""
プレゼンテーションモード制御（細分化された機能制御）
UIでは一括設定として「通知やスリープ」だが、内部的には以下を個別制御可能：
- disable_screensaver: スクリーンセーバーを無効化
- prevent_sleep: システムスリープを防止  
- block_notifications: 通知をブロック
"""
import ctypes
from ctypes import wintypes
import logging
import os
import sys

# Windows API定数
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
ES_AWAYMODE_REQUIRED = 0x00000040

# PowerRequestType列挙体
PowerRequestDisplayRequired = 0
PowerRequestSystemRequired = 1
PowerRequestAwayModeRequired = 2


class PresentationModeController:
    """プレゼンテーションモード制御クラス（細分化された機能制御）"""

    def __init__(self, silent_mode=False, config=None):
        self.silent_mode = silent_mode  # 通知をオフにするかどうか

        # デフォルト設定（UIからは一括でON/OFFされるが、内部では個別制御可能）
        default_config = {
            'disable_screensaver': True,     # スクリーンセーバーを無効化
            'prevent_sleep': True,           # システムスリープを防止
            'block_notifications': False,    # 通知をブロック（現在未実装）
        }

        # 設定をマージ（設定ファイルからの高度な制御も可能）
        self.features = default_config.copy()
        if config and isinstance(config, dict):
            self.features.update(config)

        self.kernel32 = ctypes.windll.kernel32
        self.powrprof = ctypes.windll.powrprof

        # SetThreadExecutionStateの定義
        self.kernel32.SetThreadExecutionState.argtypes = [wintypes.DWORD]
        self.kernel32.SetThreadExecutionState.restype = wintypes.DWORD

        # PowerCreateRequestの定義
        try:
            self.powrprof.PowerCreateRequest.argtypes = [
                ctypes.POINTER(wintypes.HANDLE)]
            self.powrprof.PowerCreateRequest.restype = wintypes.HANDLE

            # PowerSetRequestの定義
            self.powrprof.PowerSetRequest.argtypes = [
                wintypes.HANDLE, wintypes.DWORD]
            self.powrprof.PowerSetRequest.restype = wintypes.BOOL

            # PowerClearRequestの定義
            self.powrprof.PowerClearRequest.argtypes = [
                wintypes.HANDLE, wintypes.DWORD]
            self.powrprof.PowerClearRequest.restype = wintypes.BOOL

            # CloseHandleの定義
            self.kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
            self.kernel32.CloseHandle.restype = wintypes.BOOL

            self.power_request_available = True
        except AttributeError:
            self.power_request_available = False
            if not self.silent_mode:
                logging.info(
                    "PowerRequest APIが利用できません。SetThreadExecutionStateを使用します。")

        self.power_request_handle = None
        self.presentation_mode_active = False

        # 通知ブロック機能用の状態管理
        self.notification_blocking_active = False
        self.saved_notification_settings = {}

    def _log_info(self, message):
        """サイレントモードを考慮したログ出力"""
        if not self.silent_mode:
            logging.info(message)

    def _log_warning(self, message):
        """サイレントモードを考慮した警告ログ出力"""
        if not self.silent_mode:
            logging.warning(message)

    def _log_error(self, message):
        """エラーログは常に出力"""
        logging.error(message)

    def update_settings(self, silent_mode=None, config=None):
        """設定を更新する（既存インスタンスの設定変更用）"""
        settings_changed = False

        # サイレントモードの更新
        if silent_mode is not None and self.silent_mode != silent_mode:
            self.silent_mode = silent_mode
            settings_changed = True

        # 機能設定の更新
        if config is not None:
            new_features = self.features.copy()
            if isinstance(config, dict):
                new_features.update(config)
                if new_features != self.features:
                    self.features = new_features
                    settings_changed = True

        if settings_changed:
            self._log_info(
                f"プレゼンテーションモード設定を更新しました (silent: {self.silent_mode}, features: {self.features})")

    def enable_presentation_mode(self):
        """プレゼンテーションモードを有効にする（設定に基づいて各機能を個別制御）"""
        try:
            if self.presentation_mode_active:
                return True

            success_count = 0
            total_features = 0

            # 1. スクリーンセーバー無効化 & スリープ防止
            if self.features.get('disable_screensaver') or self.features.get('prevent_sleep'):
                total_features += 1
                if self._enable_power_management():
                    success_count += 1
                    self._log_info("スクリーンセーバー無効化・スリープ防止を有効にしました")

            # 2. 通知ブロック（将来実装予定）
            if self.features.get('block_notifications'):
                total_features += 1
                if self._enable_notification_blocking():
                    success_count += 1
                    self._log_info("通知ブロックを有効にしました")
                else:
                    self._log_warning("通知ブロックは現在未実装です")

            # 成功判定（有効な機能の半分以上が成功すれば成功とみなす）
            success = success_count > 0 and (
                total_features == 0 or success_count >= total_features * 0.5)
            self.presentation_mode_active = success

            if success:
                active_features = []
                if self.features.get('disable_screensaver'):
                    active_features.append("スクリーンセーバー無効化")
                if self.features.get('prevent_sleep'):
                    active_features.append("スリープ防止")
                if self.features.get('block_notifications'):
                    active_features.append("通知ブロック")

                self._log_info(
                    f"プレゼンテーションモードを有効にしました（機能: {', '.join(active_features)}）")

            return success

        except Exception as e:
            self._log_error(f"プレゼンテーション設定有効化エラー: {e}")
            return False

    def _enable_power_management(self):
        """電源管理制御（スクリーンセーバー無効化・スリープ防止）"""
        try:
            success = False

            # 1. PowerRequest API（Windows Vista以降）
            if self.power_request_available:
                try:
                    # REASON_CONTEXT構造体を作成
                    reason_context = wintypes.HANDLE()

                    # PowerCreateRequestを呼び出し
                    self.power_request_handle = self.powrprof.PowerCreateRequest(
                        ctypes.byref(reason_context)
                    )

                    if self.power_request_handle and self.power_request_handle != -1:
                        # 必要な機能に基づいてリクエストを設定
                        requests_success = []

                        if self.features.get('disable_screensaver'):
                            display_success = self.powrprof.PowerSetRequest(
                                self.power_request_handle,
                                PowerRequestDisplayRequired
                            )
                            requests_success.append(display_success)

                        if self.features.get('prevent_sleep'):
                            system_success = self.powrprof.PowerSetRequest(
                                self.power_request_handle,
                                PowerRequestSystemRequired
                            )
                            requests_success.append(system_success)

                        if all(requests_success) and len(requests_success) > 0:
                            success = True
                            self._log_info("PowerRequest APIで電源管理制御を有効にしました")
                        else:
                            self._log_warning("PowerSetRequestが失敗しました")
                    else:
                        self._log_warning("PowerCreateRequestが失敗しました")

                except Exception as e:
                    self._log_warning(f"PowerRequest API使用中にエラー: {e}")

            # 2. SetThreadExecutionState（フォールバック/メイン）
            if not success:
                flags = ES_CONTINUOUS
                if self.features.get('prevent_sleep'):
                    flags |= ES_SYSTEM_REQUIRED
                if self.features.get('disable_screensaver'):
                    flags |= ES_DISPLAY_REQUIRED

                result = self.kernel32.SetThreadExecutionState(flags)
                if result != 0:
                    success = True
                    self._log_info("SetThreadExecutionStateで電源管理制御を有効にしました")
                else:
                    self._log_error("SetThreadExecutionStateが失敗しました")

            return success

        except Exception as e:
            self._log_error(f"電源管理制御エラー: {e}")
            return False

    def _enable_notification_blocking(self):
        """通知ブロック機能（Windows 10以降のFocus Assist APIを使用）"""
        try:
            # Windows 10以降のFocus Assist機能を制御
            # レジストリ経由でQuiet Hours（Focus Assist）を有効化
            import winreg

            try:
                # 現在の設定を保存（後で復元するため）
                self._save_current_focus_assist_state()

                # Focus Assistの設定を変更
                key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings"

                # システム全体の通知を一時的に無効化
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    # Focus Assistの優先度モードを設定（2 = アラームのみ許可）
                    winreg.SetValueEx(
                        key, "NOC_GLOBAL_SETTING_ALLOW_NOTIFICATION_SOUND", 0, winreg.REG_DWORD, 0)
                    winreg.SetValueEx(
                        key, "NOC_GLOBAL_SETTING_ALLOW_CRITICAL_TOASTS", 0, winreg.REG_DWORD, 1)
                    winreg.SetValueEx(
                        key, "NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK", 0, winreg.REG_DWORD, 0)

                # Windows Notification Manager APIを使用した制御も試行
                try:
                    # より直接的な通知制御
                    import ctypes.wintypes
                    user32 = ctypes.windll.user32

                    # システム通知を一時的に無効化するためのパラメータ設定
                    # SPI_SETMESSAGEDURATION を使用して通知の表示時間を最小化
                    user32.SystemParametersInfoW(
                        0x2017, 0, ctypes.c_int(1), 0)  # 通知時間を最小に

                except Exception as e:
                    self._log_warning(f"システム通知制御での警告: {e}")

                self.notification_blocking_active = True
                self._log_info("通知ブロック機能を有効にしました（Focus Assist使用）")
                return True

            except PermissionError:
                self._log_warning("通知ブロック: レジストリへのアクセス権限がありません")
                return False
            except FileNotFoundError:
                self._log_warning("通知ブロック: Windows 10以降の機能のため、対応していない可能性があります")
                return False

        except ImportError:
            self._log_warning("通知ブロック: winregモジュールが利用できません")
            return False
        except Exception as e:
            self._log_error(f"通知ブロック機能エラー: {e}")
            return False

    def _save_current_focus_assist_state(self):
        """現在のFocus Assist設定を保存（後で復元するため）"""
        try:
            import winreg
            self.saved_notification_settings = {}

            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings"

            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
                    settings_to_save = [
                        "NOC_GLOBAL_SETTING_ALLOW_NOTIFICATION_SOUND",
                        "NOC_GLOBAL_SETTING_ALLOW_CRITICAL_TOASTS",
                        "NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK"
                    ]

                    for setting in settings_to_save:
                        try:
                            value, _ = winreg.QueryValueEx(key, setting)
                            self.saved_notification_settings[setting] = value
                        except FileNotFoundError:
                            # 設定が存在しない場合はデフォルト値を設定
                            self.saved_notification_settings[setting] = 1

            except FileNotFoundError:
                # レジストリキーが存在しない場合
                pass

        except Exception as e:
            self._log_warning(f"通知設定の保存中にエラー: {e}")
            self.saved_notification_settings = {}

    def _enable_wallpaper_replacement(self):
        """デスクトップ壁紙置換機能（将来実装予定）"""
        # TODO: SystemParametersInfo APIで壁紙を一時的に変更
        # 現在は未実装
        return False

    def disable_presentation_mode(self):
        """プレゼンテーションモードを無効にする（設定に基づいて各機能を個別無効化）"""
        try:
            if not self.presentation_mode_active:
                return True

            success_count = 0
            total_features = 0

            # 1. 電源管理制御を無効化
            if self.features.get('disable_screensaver') or self.features.get('prevent_sleep'):
                total_features += 1
                if self._disable_power_management():
                    success_count += 1
                    self._log_info("電源管理制御を無効にしました")

            # 2. 通知ブロック無効化（将来実装予定）
            if self.features.get('block_notifications'):
                total_features += 1
                if self._disable_notification_blocking():
                    success_count += 1
                    self._log_info("通知ブロックを無効にしました")

            # 成功判定
            success = success_count > 0 and (
                total_features == 0 or success_count >= total_features * 0.5)
            self.presentation_mode_active = False

            if success:
                self._log_info("プレゼンテーションモードを無効にしました（通常の動作に復帰）")

            return success

        except Exception as e:
            self._log_error(f"プレゼンテーション設定無効化エラー: {e}")
            return False

    def _disable_power_management(self):
        """電源管理制御を無効化"""
        try:
            success = False

            # 1. PowerRequest APIを使用している場合
            if self.power_request_handle and self.power_request_available:
                try:
                    # リクエストをクリア
                    clear_results = []

                    if self.features.get('disable_screensaver'):
                        display_clear = self.powrprof.PowerClearRequest(
                            self.power_request_handle,
                            PowerRequestDisplayRequired
                        )
                        clear_results.append(display_clear)

                    if self.features.get('prevent_sleep'):
                        system_clear = self.powrprof.PowerClearRequest(
                            self.power_request_handle,
                            PowerRequestSystemRequired
                        )
                        clear_results.append(system_clear)

                    # ハンドルをクローズ
                    handle_closed = self.kernel32.CloseHandle(
                        self.power_request_handle)

                    if all(clear_results) and handle_closed and len(clear_results) > 0:
                        success = True
                        self._log_info("PowerRequest APIで電源管理制御を無効にしました")

                    self.power_request_handle = None

                except Exception as e:
                    self._log_warning(f"PowerRequest API無効化中にエラー: {e}")

            # 2. SetThreadExecutionStateで無効化
            if not success or not self.power_request_available:
                result = self.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
                if result != 0:
                    success = True
                    self._log_info("SetThreadExecutionStateで電源管理制御を無効にしました")
                else:
                    self._log_warning("SetThreadExecutionState無効化が失敗しました")

            return success

        except Exception as e:
            self._log_error(f"電源管理制御無効化エラー: {e}")
            return False

    def _disable_notification_blocking(self):
        """通知ブロック無効化（保存された設定を復元）"""
        try:
            if not hasattr(self, 'notification_blocking_active') or not self.notification_blocking_active:
                return True

            import winreg

            # 保存された設定を復元
            if hasattr(self, 'saved_notification_settings') and self.saved_notification_settings:
                key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings"

                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        for setting_name, setting_value in self.saved_notification_settings.items():
                            winreg.SetValueEx(
                                key, setting_name, 0, winreg.REG_DWORD, setting_value)

                    # システム通知設定も復元
                    try:
                        import ctypes.wintypes
                        user32 = ctypes.windll.user32
                        # 通知時間設定を標準に戻す
                        user32.SystemParametersInfoW(
                            0x2017, 0, ctypes.c_int(5), 0)  # 標準の通知時間に戻す
                    except Exception as e:
                        self._log_warning(f"システム通知設定復元での警告: {e}")

                    self.notification_blocking_active = False
                    self._log_info("通知ブロック機能を無効にしました（設定を復元）")
                    return True

                except PermissionError:
                    self._log_warning("通知ブロック解除: レジストリへのアクセス権限がありません")
                    return False
                except Exception as e:
                    self._log_warning(f"通知設定復元中にエラー: {e}")
                    return False
            else:
                # 保存された設定がない場合はデフォルト値に戻す
                self._log_info("保存された通知設定がないため、デフォルト設定に復元")
                self.notification_blocking_active = False
                return True

        except ImportError:
            self._log_warning("通知ブロック解除: winregモジュールが利用できません")
            return False
        except Exception as e:
            self._log_error(f"通知ブロック解除エラー: {e}")
            return False

    def _disable_wallpaper_replacement(self):
        """デスクトップ壁紙復元（将来実装予定）"""
        # TODO: 元の壁紙に復元
        return False

    def is_presentation_mode_active(self):
        """プレゼンテーションモードがアクティブかどうかを返す"""
        return self.presentation_mode_active

    def __del__(self):
        """デストラクタ - リソースをクリーンアップ"""
        try:
            if self.presentation_mode_active:
                self.disable_presentation_mode()
        except:
            pass


# グローバルインスタンス
_presentation_controller = None


def get_presentation_controller(silent_mode=False, features_config=None):
    """プレゼンテーションモードコントローラーのシングルトンインスタンスを取得"""
    global _presentation_controller

    if _presentation_controller is None:
        # 初回作成
        _presentation_controller = PresentationModeController(
            silent_mode=silent_mode, config=features_config)
    else:
        # 既存インスタンスの設定を更新
        _presentation_controller.update_settings(
            silent_mode=silent_mode, config=features_config)

    return _presentation_controller


def enable_presentation_mode():
    """プレゼンテーションモードを有効にする（簡易インターフェース）"""
    controller = get_presentation_controller()
    return controller.enable_presentation_mode()


def disable_presentation_mode():
    """プレゼンテーションモードを無効にする（簡易インターフェース）"""
    controller = get_presentation_controller()
    return controller.disable_presentation_mode()


def is_presentation_mode_active():
    """プレゼンテーションモードがアクティブかどうかを確認（簡易インターフェース）"""
    controller = get_presentation_controller()
    return controller.is_presentation_mode_active()


def set_presentation_features(features_config):
    """プレゼンテーション機能の詳細設定（高度な設定用）

    Args:
        features_config (dict): 機能設定
            - disable_screensaver (bool): スクリーンセーバーを無効化
            - prevent_sleep (bool): システムスリープを防止
            - block_notifications (bool): 通知をブロック
    """
    global _presentation_controller
    if _presentation_controller is not None:
        _presentation_controller.features.update(features_config)
        return True
    return False


if __name__ == "__main__":
    # テスト用
    import time

    logging.basicConfig(level=logging.INFO)

    print("=== 細分化されたプレゼンテーションモード機能テスト ===")
    print("内部的に各機能を個別制御可能、UIでは一括設定として使用")

    # デフォルト設定でテスト
    print("\n--- デフォルト設定テスト ---")
    controller = get_presentation_controller()
    print(f"有効機能: {controller.features}")

    # 現在の状態を確認
    print(f"\nテスト前の状態:")
    print(
        f"  プレゼンテーションモード: {'有効' if controller.presentation_mode_active else '無効'}")

    # プレゼンテーションモードを有効にする
    print("\nプレゼンテーションモードを有効にします...")
    if enable_presentation_mode():
        print("✓ プレゼンテーションモード有効")
        print(
            f"  現在の状態: {'有効' if controller.presentation_mode_active else '無効'}")

        print("\n5秒間待機中...")
        time.sleep(5)

        print("\nプレゼンテーションモードを無効にします...")
        if disable_presentation_mode():
            print("✓ プレゼンテーションモード無効")
            print(
                f"  現在の状態: {'有効' if controller.presentation_mode_active else '無効'}")
        else:
            print("✗ プレゼンテーションモード無効化失敗")
    else:
        print("✗ プレゼンテーションモード有効化失敗")

    # カスタム設定でテスト
    print("\n--- カスタム設定テスト ---")
    custom_features = {
        'disable_screensaver': True,
        'prevent_sleep': False,  # スリープ防止は無効
        'block_notifications': True,  # 通知ブロック有効（未実装だがテスト）
    }

    set_presentation_features(custom_features)
    print(f"カスタム機能設定: {controller.features}")

    print("\nカスタム設定でプレゼンテーションモードを有効にします...")
    if enable_presentation_mode():
        print("✓ カスタム設定で有効化完了")
        time.sleep(3)
        if disable_presentation_mode():
            print("✓ カスタム設定で無効化完了")

    print("\nテスト完了")
