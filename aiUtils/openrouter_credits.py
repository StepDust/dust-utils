import requests
import json
from datetime import datetime


class OpenRouterCredits:
    # 单例实例
    _instance = None
    # 初始化标记
    _initialized = False

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, token):
        # 确保init只执行一次
        if not OpenRouterCredits._initialized:
            self.token = token
            self.send_msg_date = None  # 上次预警发送时间
            self.credits = None  # 余额对象
            OpenRouterCredits._initialized = True

            self.get_credits()

    def get_credits(self):
        # 检查是否需要重置发送日期
        if (
            self.send_msg_date is None
            or self.send_msg_date != datetime.now().date()
            or self.credits is None
        ):
            url = "https://openrouter.ai/api/v1/credits"
            headers = {"Authorization": f"Bearer {self.token}"}
            # {"data":{"total_credits":5,"total_usage":3.9272644175}}
            resp = requests.get(url, headers=headers)
            if resp.status_code == 200:
                self.credits = json.loads(resp.text)["data"]

                balance = self.credits["total_credits"] - self.credits["total_usage"]
                self.credits["balance"] = balance
                self.send_msg_date = datetime.now().date()
                return self.credits
            else:
                print("查询失败：", resp.status_code, resp.text)
                return None

        return self.credits
