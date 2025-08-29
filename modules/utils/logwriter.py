import os
import logging
from datetime import datetime
import sys


def get_app_data_dir():
    """アプリケーション用のデータディレクトリを取得"""
    try:
        # PyInstallerの場合、実行ファイルのディレクトリを取得
        if hasattr(sys, '_MEIPASS'):
            # 実行ファイルのディレクトリを取得
            exe_dir = os.path.dirname(sys.executable)
            return exe_dir
        else:
            # 開発環境では現在のディレクトリ
            return os.path.abspath(".")
    except Exception:
        return os.path.abspath(".")


def get_log_directory():
    """ログディレクトリのパスを取得"""
    app_dir = get_app_data_dir()
    log_dir = os.path.join(app_dir, 'logs')

    # ログディレクトリに書き込み権限がない場合はユーザーディレクトリを使用
    try:
        # テスト用のファイル作成を試行
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        test_file = os.path.join(log_dir, 'test_write.tmp')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)

        return log_dir
    except (OSError, PermissionError):
        # 権限がない場合はユーザーのアプリケーションデータフォルダを使用
        try:
            user_data_dir = os.path.expanduser('~')
            fallback_log_dir = os.path.join(
                user_data_dir, 'FullScreenCover', 'logs')

            if not os.path.exists(fallback_log_dir):
                os.makedirs(fallback_log_dir)

            return fallback_log_dir
        except Exception:
            # 最終フォールバック: 一時ディレクトリ
            import tempfile
            temp_log_dir = os.path.join(
                tempfile.gettempdir(), 'FullScreenCover', 'logs')
            if not os.path.exists(temp_log_dir):
                os.makedirs(temp_log_dir)
            return temp_log_dir


class FullScreenCoverLogger:
    """FullScreenCover専用ログシステム"""

    def __init__(self, log_dir=None, enable_console=True, enable_file=True):
        self.log_dir = log_dir if log_dir else get_log_directory()
        self.enable_console = enable_console
        self.enable_file = enable_file
        self.logger = None
        self._setup_logging()

    def _setup_logging(self):
        """ログシステムを設定"""
        # ログディレクトリの作成
        if self.enable_file and not os.path.exists(self.log_dir):
            try:
                os.makedirs(self.log_dir)
            except Exception as e:
                print(f"ログディレクトリ作成エラー: {e}")
                self.enable_file = False

        # ロガーの作成
        self.logger = logging.getLogger('FullScreenCover')
        self.logger.setLevel(logging.DEBUG)

        # 既存のハンドラーをクリア（重複防止）
        self.logger.handlers.clear()

        # フォーマッターの設定
        formatter = logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # ファイルハンドラーの設定
        if self.enable_file:
            try:
                log_filename = os.path.join(
                    self.log_dir, f'{datetime.now():%Y-%m-%d}.log')
                file_handler = logging.FileHandler(
                    log_filename, encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)

                # ログディレクトリの場所を記録
                self.logger.info(f"ログファイル: {log_filename}")
            except Exception as e:
                print(f"ログファイル作成エラー: {e}")
                self.enable_file = False

        # コンソールハンドラーの設定
        if self.enable_console:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

    def debug(self, message):
        """デバッグレベルのログ"""
        if self.logger:
            self.logger.debug(str(message))

    def info(self, message):
        """情報レベルのログ"""
        if self.logger:
            self.logger.info(str(message))

    def warning(self, message):
        """警告レベルのログ"""
        if self.logger:
            self.logger.warning(str(message))

    def error(self, message):
        """エラーレベルのログ"""
        if self.logger:
            self.logger.error(str(message))

    def critical(self, message):
        """致命的エラーレベルのログ"""
        if self.logger:
            self.logger.critical(str(message))

    def get_log_path(self):
        """現在のログファイルのパスを取得"""
        if self.enable_file:
            return os.path.join(self.log_dir, f'{datetime.now():%Y-%m-%d}.log')
        return None


# グローバルロガーインスタンス
_global_logger = None


def get_logger(log_dir=None, enable_console=True, enable_file=True):
    """グローバルロガーインスタンスを取得"""
    global _global_logger
    if _global_logger is None:
        _global_logger = FullScreenCoverLogger(
            log_dir, enable_console, enable_file)
    return _global_logger


def setup_logging(log_dir=None, enable_console=True, enable_file=True):
    """ログシステムを初期化（後方互換性のため）"""
    return get_logger(log_dir, enable_console, enable_file)


def get_current_log_path():
    """現在のログファイルのパスを取得"""
    logger = get_logger()
    return logger.get_log_path()


# 便利関数
def log_debug(message):
    """デバッグログの便利関数"""
    logger = get_logger()
    logger.debug(message)


def log_info(message):
    """情報ログの便利関数"""
    logger = get_logger()
    logger.info(message)


def log_warning(message):
    """警告ログの便利関数"""
    logger = get_logger()
    logger.warning(message)


def log_error(message):
    """エラーログの便利関数"""
    logger = get_logger()
    logger.error(message)


def log_critical(message):
    """致命的エラーログの便利関数"""
    logger = get_logger()
    logger.critical(message)
