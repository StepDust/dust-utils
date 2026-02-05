import requests
import json
import time
import re
import hmac
import hashlib
import base64
import urllib.parse
import logging
from .openrouter_credits import OpenRouterCredits

# 创建模块专用记录器
logger = logging.getLogger(__name__)


class AIChat:
    """
    这是一个AI聊天工具类，主要功能包括:
    1. 与AI大模型进行对话交互，支持文本和图像输入
    2. 支持多种AI模型配置和服务商(OpenAI、OpenRouter等)
    3. 记录对话token使用量、费用和响应时间
    4. 提供JSON和代码格式修复功能
    5. 提供余额预警功能，支持钉钉通知

    主要方法:
    - send_message(): 发送消息并获取AI响应，支持文本和图像输入
    - clear_message(): 清空对话历史
    - fix_json(): 修复不规范的JSON字符串
    - fix_code(): 移除代码块标记，支持多种编程语言
    - check_credits(): 查询账户余额
    - send_dingtalk_message(): 发送钉钉预警消息
    """

    def __init__(self, config):
        """
        初始化AI聊天实例

        Args:
            config: 包含AI配置信息的字典，需要包含hostsUrl、apiKey和model字段
        """
        try:
            from openai import OpenAI

            self.openai = OpenAI
        except ImportError:
            raise ImportError(
                "检测到未安装 openai。请执行 'pip install openai' 以使用此功能。"
            )

        self.base_url = config.get("baseUrl")
        self.api_key = config.get("apiKey")
        self.model = config.get("model")
        self.mask = config.get("mask")
        self.modelType = config.get("modelType")

        # 初始化ai角色定义
        self.messageList = [
            {
                "role": "system",
                "content": self.mask,
            }
        ]

        # 金额定价
        self.input_price = config.get("inputPrice", 0) / 1000  # 输入金额定价
        self.output_price = config.get("outputPrice", 0) / 1000  # 输出金额定价
        self.price = 0  # 已使用总金额
        self.useToken = 0  # 已使用总token
        self.useTime = 0  # 已使用总时间

        # 其他信息
        self.credits = None
        self.creditAlert = config.get("creditAlert", 0)
        self.sendCount = 0  # 发送次数

        # 查询余额
        self.check_credits()

    def send_message(self, message, image_url=None):
        """
        发送消息到AI服务并获取响应

        Args:
            message: 要发送给AI的消息内容

        Returns:
            str: AI的响应消息
        """
        try:

            client = self.openai(
                # 若没有配置环境变量,请用阿里云百炼API Key将下行替换为:api_key="sk-xxx",
                api_key=self.api_key,
                base_url=self.base_url,
            )

            print("")
            logger.info(f"{message}", extra={"color": "#31bdec"})
            # 发送对话请求
            self.messageList.append({"role": "user", "content": message})

            # 记录开始时间
            start_time = time.time()
            assistant_output = client.chat.completions.create(
                model=self.model,
                messages=self.messageList,
                extra_body={
                    "enable_thinking": False  # 添加此参数，在非流式调用中禁用深度思考功能
                },
            )

            # 计算响应时间
            response_time = time.time() - start_time
            self.useTime += response_time  # 累计使用时间

            # 获取实际的回复内容
            response_content = assistant_output.choices[0].message.content

            # 计算本次对话的token使用量和金额
            input_token = assistant_output.usage.prompt_tokens
            output_token = assistant_output.usage.completion_tokens
            self.useToken += input_token + output_token  # 累计使用token
            self.price += (
                input_token * self.input_price + output_token * self.output_price
            )  # 累计使用金额

            # 将大模型的回复信息添加到对话列表中
            self.messageList.append({"role": "assistant", "content": response_content})

            logger.info(response_content + "")
            # 输出黄色的token使用量和本次对话金额
            logger.info(
                f"使用Token: {input_token + output_token}\t金额: {(input_token * self.input_price + output_token * self.output_price):.6f}元\t响应时间: {response_time:.2f}秒\tAI模型: {self.model}\tbaseURL: {self.base_url}",
                extra={"color": "#ffb800"},
            )

            self.sendCount += 1  # 发送次数加1
            return response_content

        except requests.exceptions.RequestException as e:
            logger.error(f"请求发生错误: {e}")
            return None
        except json.JSONDecodeError as e:
            logger.error(f"响应解析错误: {e}")
            return None
        except Exception as e:
            logger.error(f"发生未知错误: {e}")
            return None

    def clear_message(self):
        """
        清空消息列表
        """
        self.messageList = [
            {
                "role": "system",
                "content": self.mask,
            }
        ]

    def fix_json(self, json_str, out_obj=True):
        """
        修复不规范的JSON字符串，支持自动修复常见的JSON格式错误

        Args:
            json_str: 可能不规范的JSON字符串
            out_obj: 是否返回Python对象，True返回dict对象，False返回JSON字符串

        Returns:
            Union[dict, str]: 根据out_obj参数返回修复后的JSON对象或字符串
            - 当out_obj=True时返回dict对象
            - 当out_obj=False时返回格式化的JSON字符串
        """
        if not json_str:
            if out_obj:
                return {}
            else:
                return "{}"

        try_count = 0
        max_try_count = 3  # 最大重试次数

        while try_count < max_try_count:

            json_str = self.fix_code(json_str, ["json"]).replace("\n", "")

            # 移除所有 <style>...</style> 内容
            json_str = re.sub(r"<style>.*?</style>", "", json_str, flags=re.DOTALL)
            # 使用正则表达式查找缺少引号的键值对
            # 匹配模式: "key":value 其中value不是以引号、数字、{、[、true、false、null开头的
            pattern = r'("[^"]+":)\s*([^\s"\d\{\[trfn][^,\}\]]*)'  # 匹配没有引号的值
            json_str = re.sub(pattern, r'\1"\2"', json_str)

            # 修复没有使用双引号包裹的属性名
            pattern_unquoted_key = r"(\{|\,)\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*:"
            json_str = re.sub(pattern_unquoted_key, r'\1"\2":', json_str)

            try:
                jsonObj = json.loads(json_str)
                if out_obj:
                    return jsonObj
                return json.dumps(jsonObj, ensure_ascii=False)
            except json.JSONDecodeError:
                try_count += 1
                jsonErrorQuestion = f"```{json_str}```这是一个json格式错误的文本，请帮我修正，请注意属性应被双引号包裹，我只要修正后的json，不要输出其他内容，也不要增删属性，保持json数据结构不变，属性值中可能存在双引号，注意转义"
                json_str = self.send_message(jsonErrorQuestion)

        # 超过最大重试次数后抛出异常
        if try_count >= max_try_count:
            error_msg = f"JSON修复失败,已重试{max_try_count}次"
            logger.error(f"{error_msg}")  # 红色打印错误信息
            raise ValueError(error_msg)

    def fix_js(self, javascript_code):
        """
        修复JavaScript代码中的语法错误

        Args:
            javascript_code: 包含JavaScript代码的字符串

        Returns:
            str: 修复后的JavaScript代码字符串
        """
        try:
            import esprima
        except ImportError:
            raise ImportError("该功能需要 esprima，请执行：pip install esprima")

        if not javascript_code:
            return ""

        # ----------- 内部工具函数 -----------
        def strip_comments(code: str) -> str:
            """去掉 JS 单行和多行注释"""
            code = re.sub(r"/\*[\s\S]*?\*/", "", code)  # 多行注释
            code = re.sub(r"//[^\n]*", "", code)  # 单行注释
            return code

        def sanitize_js(code: str) -> str:
            """修复字符串字面量中被意外打断的换行，替换成 '\\n'"""
            code = re.sub(r"'[\r\n]+'", r"'\\n'", code)  # 单引号里的非法换行
            code = re.sub(r'"[\r\n]+"', r'"\\n"', code)  # 双引号里的非法换行
            return code

        def js_syntax_ok(code: str) -> bool:
            """仅做语法检查，返回 True/False"""
            try:
                esprima.parseScript(code, tolerant=False)
                return True
            except esprima.Error:
                return False

        # ----------- 内部工具函数结束 -----------

        try_count = 0
        max_try_count = 3

        while try_count < max_try_count:
            javascript_code = self.fix_code(javascript_code)  # 移除代码块标记
            no_comment_code = strip_comments(javascript_code)  # 1. 去注释
            sanitized_code = sanitize_js(no_comment_code)  # 2. 修非法换行

            if js_syntax_ok(sanitized_code):  # 3. 语法校验
                return javascript_code

            # 语法仍报错 → 交给 AI 修复
            try_count += 1
            js_error_question = (
                f"```{javascript_code}```\n"
                f"这是一个 JavaScript 代码，其中可能存在语法错误，请帮我修正。"
                f"我只要修正后的代码，不要输出其他内容，也不要改变代码逻辑或者修改变量、属性名称以及对应值。"
            )
            javascript_code = self.send_message(js_error_question)

        # 超过最大重试次数
        error_msg = f"JavaScript 代码修复失败，已重试 {max_try_count} 次"
        logger.error(error_msg)
        raise ValueError(error_msg)

    def fix_mermaid(self, mermaid_code):
        """
        修复Mermaid图表代码中的语法错误

        Args:
            mermaid_code: 包含Mermaid图表代码的字符串

        Returns:
            str: 修复后的Mermaid图表代码字符串
        """
        try:
            import mermaid as md
        except ImportError:
            raise ImportError(
                "生成 Mermaid 图表需要 mermaid，请执行：pip install mermaid-py"
            )

        if not mermaid_code:
            return ""

        try_count = 0
        max_try_count = 3  # 最大重试次数

        # while try_count < max_try_count:
        #     # 移除代码块标记
        #     mermaid_code = self.fix_code(mermaid_code, ["mermaid"])

        #     try:
        #         code = mermaid_code.replace("\\n", "\n")
        #         # 使用pymermaid检查Mermaid语法
        #         mermaid = md.Mermaid(code)
        #         if mermaid.svg_response.status_code != 200:
        #             raise ValueError(f"mermaid字符串异常:{mermaid_code}")
        #         return mermaid_code
        #     except Exception as e:
        #         try_count += 1
        #         # 发送修复请求给AI
        #         mermaid_error_question = f"```{mermaid_code}```这是一个Mermaid图表代码，其中可能存在语法错误，请帮我修正，我只要修正后的代码，不要输出其他内容，也不要改变图表逻辑或者修改节点、关系以及对应的描述"
        #         mermaid_code = self.send_message(mermaid_error_question)

        # # 超过最大重试次数后抛出异常
        # if try_count >= max_try_count:
        #     error_msg = f"Mermaid图表代码修复失败,已重试{max_try_count}次"
        #     logger.error(f"{error_msg}")
        #     raise ValueError(error_msg)

    def fix_code(self, code, additional_tags=[]):
        """
        移除代码字符串中的代码块标记（如```python等）

        Args:
            code: 需要处理的代码字符串，可能包含代码块标记
            additional_tags: 额外的编程语言标签列表，用于扩展默认支持的语言类型

        Returns:
            str: 移除代码块标记后的代码字符串，保持代码内容不变
        """
        # 定义常见编程语言列表
        languages = [
            # 后端语言
            "python",
            "java",
            "c",
            "c++",
            "c#",
            "csharp",
            "go",
            "rust",
            "php",
            "ruby",
            "kotlin",
            "scala",
            "perl",
            "r",
            # 前端语言
            "javascript",
            "typescript",
            "html",
            "css",
            "sass",
            "less",
            "vue",
            "react",
            "angular",
            # 数据库
            "sql",
            "mysql",
            "postgresql",
            "mongodb",
            # 标记语言
            "xml",
            "yaml",
            "json",
            "markdown",
            # 脚本语言
            "shell",
            "bash",
            "powershell",
            "batch",
            # 移动开发
            "swift",
            "objective-c",
            "dart",
            "flutter",
            # 其他语言
            "matlab",
            "assembly",
            "fortran",
            "cobol",
            "pascal",
            "ada",
            "lisp",
            "prolog",
            "haskell",
            "erlang",
            "elixir",
            "lua",
        ]

        if additional_tags:
            languages.extend(additional_tags)

        if not code:
            return ""

        # 使用正则表达式移除所有语言的代码块标记
        for lang in languages:
            # 使用re.escape转义语言名，避免元字符引发正则错误
            pattern = re.compile(rf"```{re.escape(lang)}[\s\n]", re.IGNORECASE)
            code = pattern.sub("", code)

        # 移除剩余的代码块标记和换行符
        code = code.replace("```", "")

        return code

    def check_credits(self):
        """
        检查当前账户余额
        """
        if self.creditAlert is None or self.creditAlert <= 0:
            return

        # OpenRouter平台余额查询
        if self.base_url and "openrouter" in self.base_url:
            # 初始化OpenRouterCredits对象
            credits = OpenRouterCredits(self.api_key)
            self.credits = credits.get_credits()
            # 检查余额是否低于预警值
            if self.credits["balance"] < self.creditAlert:
                # 发送钉钉预警消息
                self.send_dingtalk_message(self.credits["balance"])

    def send_dingtalk_message(self, balance):
        """发送钉钉预警消息"""
        try:
            # 钉钉机器人webhook地址
            webhook = "https://oapi.dingtalk.com/robot/send?access_token=20eb73ffefa3c10564d57301297a6cbb3012f0772d051d5f368102b1fd4c3a45"
            # 钉钉机器人密钥
            secret = (
                "SEC95d2a74bda471c22b330199caead52a227a8ca622d84fc968b21df2e07e2cde9"
            )

            def get_timestamp_and_sign(secret):
                timestamp = str(round(time.time() * 1000))
                string_to_sign = f"{timestamp}\n{secret}"
                hmac_code = hmac.new(
                    secret.encode("utf-8"),
                    string_to_sign.encode("utf-8"),
                    digestmod=hashlib.sha256,
                ).digest()
                sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))
                return timestamp, sign

            timestamp, sign = get_timestamp_and_sign(secret)
            webhook_url = f"{webhook}&timestamp={timestamp}&sign={sign}"

            # 消息内容
            message = {
                "msgtype": "text",
                "text": {
                    "content": f"⚠️ OpenRouter API 余额预警\n⚠️ 当前余额: {balance:.2f} 美元\n⚠️ 预警余额: {self.creditAlert} 美元\n🪙 充值地址：https://openrouter.ai/settings/credits"
                },
            }

            # 发送POST请求
            headers = {"Content-Type": "application/json"}
            response = requests.post(
                webhook_url, headers=headers, data=json.dumps(message)
            )

            if response.status_code == 200:
                logger.success("钉钉预警消息发送成功")
            else:
                logger.error(
                    f"钉钉消息发送失败: {response.status_code} - {response.text}"
                )

        except Exception as e:
            logger.error(f"发送钉钉消息时发生错误: {str(e)}")
