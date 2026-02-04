# 个人公共工具库 dust_utils

## 更新与安装方式

### 更新 common_utils 库本身（在 common_utils 目录下操作）
- **修改 .py 文件、docstring 等**：保存后立即在所有项目中生效，无需重新安装。
- **重新安装/更新**（仅在需要时执行）：
  ```bash
  pip install -e .                  # 基础安装
  ```
- **修改 pyproject.toml** 后（如新增或调整依赖）：
需要在每个使用该功能的其他项目中，重新运行对应的安装命令

## 虚拟环境推荐使用方式（强烈建议）
每个项目应使用独立的虚拟环境，以避免依赖冲突和版本问题。
- 创建虚拟环境
    ```bash
    python -m venv venv 
    ```
- 激活虚拟环境
    - Windows
        ```bash
        venv\Scripts\activate
        ```
    - macOS / Linux
        ```bash
        source venv/bin/activate
        ```
- 退出虚拟环境
    ```bash
    deactivate
    ```
## 安装 common_utils 库
- 基础安装
    ```bash
    pip install -e ../common_utils
    ```
- 安装可选功能
    - AI 功能
        ```bash
        pip install -e ../common_utils[ai]
        ```
    - wx和AI功能
        ```bash
        pip install -e ../common_utils[wx,ai]
        ```
>  <br>
> ./common_utils，只是一个指向 common_utils 目录的路径
> <br><br>

## 安装其他项目依赖
```bash
pip install -r requirements.txt
```

## 输出解释器路径
```bash
python -c "import sys; print(sys.executable)"
```
## 输出python版本&位数
```bash
python -c "import sys; print(sys.version)"
```

## 注意事项
- 修改 common_utils 源码后，所有项目自动使用最新代码（得益于 editable 模式）。
- 只有 pyproject.toml 的依赖发生变化时，才需在相关项目中重新运行安装命令。
- 建议按需导入模块，例如：
```python
from common_utils.logger_setup import setup_logger
from common_utils.aiUtils.ai_chat import AIChat
```
这样可以避免加载不必要的第三方库。


## 常用模块
- **pipreqs**
   - 用于自动生成 requirements.txt 文件，包含项目所有依赖（包括直接和间接依赖）。
   - 安装：`pip install pipreqs`
   - 使用：在项目根目录下运行 `pipreqs . --encoding=utf-8 --ignore=__pycache__,venv --force`
   - 使用utf8编码，忽略 __pycache__ 和 venv 目录，强制覆盖已有的 requirements.txt 文件。
- **pyinstaller**
   - 用于将 Python 脚本打包成独立的可执行文件，方便在无 Python 环境的机器上运行。
   - 安装：`pip install pyinstaller`
   - 使用：在项目根目录下运行 `pyinstaller --onefile your_script.py`（将生成 dist 目录，包含可执行文件）
- **requests**
   - 用于发送 HTTP 请求，处理响应等。
   - 安装：`pip install requests`
   - 使用：在 Python 脚本中导入 `import requests` 即可开始使用。
- **colorama**
   - 用于在终端中输出彩色文本，增强可读性。
   - 安装：`pip install colorama`
   - 使用：在 Python 脚本中导入 `import colorama` 即可开始使用。
- **wcwidth**
   - 用于处理终端中显示的宽字符宽度，避免显示异常。
   - 安装：`pip install wcwidth`
   - 使用：在 Python 脚本中导入 `import wcwidth` 即可开始使用。
- **wxPython**
   - 用于创建图形用户界面（GUI）应用程序。
   - 安装：`pip install wxPython`
   - 使用：在 Python 脚本中导入 `import wx` 即可开始使用。
- **playwright**
   - 用于自动化测试和浏览器操作。
   - 安装：`pip install playwright`
   - 使用：在 Python 脚本中导入 `import playwright` 即可开始使用。

## spec移除打包依赖
- pywin32
- Pillow
- mermaid-py
- esprima
- openai
- wxPython
- pymysql
- playwright



