import ctypes
import os


def is_subpath(child: str, parent: str) -> bool:
    """Return True if 'child' is the same as or inside 'parent'."""
    child_abs = os.path.abspath(child)
    parent_abs = os.path.abspath(parent)
    return os.path.commonpath([child_abs, parent_abs]) == parent_abs


# Windows用 無操作時間取得


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]


def get_idle_duration():
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(lii)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)) == 0:
        return 0
    millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
    return millis / 1000.0  # 秒で返す
