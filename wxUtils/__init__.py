try:
    import wx
except ImportError:
    raise RuntimeError("wxUtils 需要 wxPython，请安装：pip install common-utils[wx]")

from .wx_utils import WxUtils
from .mini_alert import MiniAlert

__all__ = [
    "WxUtils",
    "MiniAlert",
]
