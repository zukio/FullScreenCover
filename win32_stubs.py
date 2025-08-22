# Pylance type checking fixes for pywin32
# This file helps Pylance recognize win32 modules

# Type stubs for win32 modules
from typing import Any, Tuple, Optional

# win32gui stubs


def GetForegroundWindow() -> int: ...
def GetWindowText(hwnd: int) -> str: ...
def GetClassName(hwnd: int) -> str: ...
def GetWindowRect(hwnd: int) -> Tuple[int, int, int, int]: ...

# win32api stubs


def GetSystemMetrics(index: int) -> int: ...


# win32con constants
SM_CXSCREEN: int = 0
SM_CYSCREEN: int = 1
