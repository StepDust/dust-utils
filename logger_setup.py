import logging
import time
import os
import sys
from typing import Any
import json
from logging.handlers import RotatingFileHandler

# 全局变量声明
log_file = None

# 定义SUCCESS日志级别
SUCCESS = 25  # 在INFO(20)和WARNING(30)之间

# 定义DIVIDER日志级别
DIVIDER = 15  # 在DEBUG(10)和INFO(20)之间

logging.addLevelName(DIVIDER, "DIVIDER")
logging.addLevelName(SUCCESS, "SUCCESS")

# 初始化颜色配置 (兼容 Windows)
try:
    import colorama

    colorama.init()
except ImportError:
    pass  # 非 Windows 系统或未安装 colorama 不影响运行


class ColorFormatter(logging.Formatter):
    """支持 16 进制颜色代码的自定义 Formatter"""

    COLOR_CODES = {
        "DIVIDER": "#eeeeee",
        "DEBUG": "#2196F3",
        "INFO": "#999999",
        "SUCCESS": "#4CAF50",
        "WARNING": "#FF9800",
        "ERROR": "#F44336",
        "CRITICAL": "#B71C1C",
    }

    def _hex_to_ansi(self, hex_color):
        hex_color = hex_color.lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return f"\033[38;2;{r};{g};{b}m"

    def format(self, record):
        default_color = self.COLOR_CODES.get(record.levelname, "#FFFFFF")
        color = getattr(record, "color", default_color)
        ansi_color = self._hex_to_ansi(color)
        message = super().format(record)
        return f"{ansi_color}{message}\033[0m"


def safe_to_dict(
    obj: Any, seen: set | None = None, max_depth: int = 3, current_depth: int = 0
) -> Any:
    """
    把对象转成可序列化的 dict，支持嵌套、列表、循环引用检测
    """
    if seen is None:
        seen = set()

    obj_id = id(obj)
    # print(f"对象：{obj}\t深度：{current_depth}")
    if obj_id in seen:
        return f"<循环引用: {type(obj).__name__}>"

    if current_depth > max_depth * 2 + 1:
        return f"<深度超出 {max_depth}>"

    seen.add(obj_id)

    if isinstance(obj, (list, tuple)):
        return [safe_to_dict(x, seen.copy(), max_depth, current_depth + 1) for x in obj]

    if isinstance(obj, dict):
        return {
            k: safe_to_dict(v, seen.copy(), max_depth, current_depth + 1)
            for k, v in obj.items()
        }

    if hasattr(obj, "__dict__") and not isinstance(obj, type):
        d = {}
        for k, v in vars(obj).items():
            d[k] = safe_to_dict(v, seen.copy(), max_depth, current_depth + 1)
        return {**d, "__class__": obj.__class__.__name__}

    if hasattr(obj, "_asdict"):  # dataclass / namedtuple
        return safe_to_dict(obj._asdict(), seen.copy(), max_depth, current_depth + 1)

    return obj


def logger_success(self, message, *args, **kwargs):
    """记录SUCCESS级别日志"""
    # stacklevel=3 的含义是：跳过当前函数、再跳过包装函数，定位到调用的代码处
    self.log(SUCCESS, message, *args, stacklevel=3, **kwargs)


def logger_divider(self, message="", max_len=50, char="=", *args, **kwargs):
    """记录 DIVIDER 分隔线日志"""
    import wcwidth

    message = message.rstrip()
    if len(message) > 0:
        message = f" {message} "

    msg_width = wcwidth.wcswidth(message)
    if msg_width >= max_len:
        self.log(DIVIDER, message, *args, stacklevel=3, **kwargs)
        return
    if msg_width >= max_len - 5:
        padding = char * (max_len - msg_width - 1)
        show_msg = f"{padding} {message}"
    else:
        left_padding = char * 5
        right_padding = char * (max_len - msg_width - 7)
        show_msg = f"{left_padding}{message}{right_padding}"
    # stacklevel=3 的含义是：跳过当前函数、再跳过包装函数，定位到调用的代码处
    self.log(DIVIDER, show_msg, *args, stacklevel=3, **kwargs)


def logger_object(
    self, object: dict | list = None, message="变量值如下：", *args, **kwargs
):
    """记录对象日志"""
    if object is None:
        self.log(logging.INFO, "这是一个空对象", *args, stacklevel=3, **kwargs)
        return
    data = safe_to_dict(object, max_depth=5)
    self.log(
        logging.INFO,
        f"{message}\n{json.dumps(data, ensure_ascii=False, indent=2, default=repr)}",
        *args,
        stacklevel=3,
        **kwargs,
    )


def logger_log_path(self):
    """返回当前 logger 的日志文件路径"""
    return log_file


def setup_logger(
    log_folder="", additional_logger_names=[], max_log_file_size=10 * 1024 * 1024
):
    """
    配置全局日志器
    :param log_folder: 日志文件夹路径
    :param additional_logger_names: 需要额外排除的日志器名称
    :param max_log_file_size: 日志文件最大大小，默认 10MB
    :return: 配置好的 logger 实例
    """
    global log_file
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    if log_folder:
        if not os.path.isabs(log_folder):
            raise ValueError("log_folder 必须是绝对路径")
        log_dir = log_folder
    else:
        current_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        log_dir = os.path.join(current_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"{timestamp}.log")
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    # 注册方法到Logger
    logging.Logger.success = logger_success
    logging.Logger.divider = logger_divider
    logging.Logger.log_path = logger_log_path
    logging.Logger.object = logger_object

    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=max_log_file_size,
        backupCount=1,
        encoding="utf-8",
    )
    file_formatter = logging.Formatter(
        "%(asctime)s [%(levelname)-7s] [%(threadName)-10s] %(filename)s:%(lineno)d - %(message)s"
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    console_handler = logging.StreamHandler()
    color_formatter = ColorFormatter(
        "%(asctime)s [%(levelname)-7s] [%(threadName)-10s] %(message)s"
    )
    console_handler.setFormatter(color_formatter)
    logger.addHandler(console_handler)
    disabled_loggers = [
        "requests",
        "urllib3",
        "chardet",
        "charset_normalizer",
        "httpcore._backends.sync",
        "httpx._client",
        "httpx",
        "openai",
        "openrouter",
        "stainless",
        "httpcore",
        "pywebview",
        "webview",
    ]
    disabled_loggers.extend(additional_logger_names)
    for logger_name in disabled_loggers:
        logging.getLogger(logger_name).setLevel(logging.ERROR)
    return logger
