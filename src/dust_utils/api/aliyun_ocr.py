# 导入必要的模块
import random  # 用于生成随机数
import requests  # 用于发送HTTP请求
import os  # 用于文件和目录操作
from urllib.parse import urlparse  # 用于URL解析
import logging  # 用于日志记录
from urllib.parse import quote, urljoin, urlunparse
import re
import json
from .api_utils import ApiUtils

# 创建模块专用记录器
logger = logging.getLogger(__name__)

# 禁用所有代理设置，确保直接连接
os.environ.pop("HTTP_PROXY", None)
os.environ.pop("HTTPS_PROXY", None)
os.environ.pop("ALL_PROXY", None)


class AliyunOCR:
    """
    阿里云OCR服务类
    服务列表：https://marketnext.console.aliyun.com/bizlist
    """

    def __init__(self):
        self.base_url = "https://gjbsb.market.alicloudapi.com"

    def post_ocrservice_advanced(self, params, app_code):
        """
        调用阿里云OCR服务的高级接口
        购买地址：https:#market.aliyun.com/detail/cmapi028554
        购买金额：0.01元/500次/月

        :param params: 请求参数
        :param app_code: 阿里云应用代码
        :return: 响应JSON数据
        """
        url = ApiUtils.url_concat(self.base_url, "/ocrservice/advanced")
        default_params = {
            "img": "",  #  图像数据：base64编码，要求base64编码后大小不超过25M，最短边至少15px，最长边最大8192px，支持jpg/png/bmp格式，和url参数只能同时存在一个
            "url": "",  # 图像url地址：图片完整URL，URL长度不超过1024字节，URL对应的图片base64编码后大小不超过25M，最短边至少15px，最长边最大8192px，支持jpg/png/bmp格式，和img参数只能同时存在一个
            "prob": False,  # 是否需要识别结果中每一行的置信度，默认不需要。 true：需要 False：不需要
            "charInfo": False,  # 是否需要单字识别功能，默认不需要。 true：需要 False：不需要
            "rotate": False,  # 是否需要自动旋转功能，默认不需要。 true：需要 False：不需要
            "table": False,  # 是否需要表格识别功能，默认不需要。 true：需要 False：不需要
            "sortPage": False,  # 字块返回顺序，False表示从左往右，从上到下的顺序，true表示从上到下，从左往右的顺序，默认False
            "noStamp": False,  # 是否需要去除印章功能，默认不需要。true：需要 False：不需要
            "figure": False,  # 是否需要图案检测功能，默认不需要。true：需要 False：不需要
            "row": False,  # 是否需要成行返回功能，默认不需要。true：需要 False：不需要
            "paragraph": False,  # 是否需要分段功能，默认不需要。true：需要 False：不需要
            "oricoord": False,  # 图片旋转后，是否需要返回原始坐标，默认不需要。true：需要  False：不需要
        }
        params = ApiUtils.combined_params(params=params, default_params=default_params)

        headers = {
            "Authorization": f"APPCODE {app_code}",
            "Content-Type": "application/json; charset=UTF-8",
        }

        # 发送POST请求并返回JSON响应
        response = requests.post(url, json=params, headers=headers, proxies=None)
        return response.json()
