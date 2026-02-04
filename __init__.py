"""
common_utils
通用工具库（按需加载）
"""

__version__ = "0.1.0"

# 只暴露“完全无重依赖”的工具
from .logger_setup import setup_logger

__all__ = [
    "setup_logger",
]
