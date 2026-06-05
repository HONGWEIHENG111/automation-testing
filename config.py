# config.py
# ==========================================
# 1. 账号与凭证信息 (Credentials & API Keys)
# ==========================================
# 网站登录账号
USERNAME = "HenryHONG"
# 网站登录密码
PASSWORD = "12345678"

# --- 多网页并行环境独立账号密码 (供 main_multiweb.py 调用) ---
USERNAME_1 = "HenryHONG"
PASSWORD_1 = "12345678"

USERNAME_2 = "weihehong3-c@my.cityu.edu.hk"  # 若环境2账号不同，可在此处单独修改
PASSWORD_2 = "1234abcdHWH"  # 若环境2密码不同，可在此处单独修改
# DeepSeek API Key (统一管理，供评价器和总结生成器调用)
DEEPSEEK_API_KEY = "sk-2aafb675ef02459595b4ea1f5e9fe040"

# ==========================================
# 2. 网址与接口 (URLs)
# ==========================================
# 测试目标网站主页
HOME_URL = "https://customs-demo.poffices.ai/"
#https://192.168.1.104/
#https://customs-demo.poffices.ai/
# DeepSeek API 请求基准地址
DEEPSEEK_BASE_URL = "https://api.deepseek.com"

# --- 多网页并行环境独立网址 ---
URL_1 = "https://customs-demo.poffices.ai/"
ENV_NAME_1 = "customs"

URL_2 = "https://poffices.ai/"
ENV_NAME_2 = "poffices"

# ==========================================
# 3. 文件路径 (Paths)
# ==========================================
# 输入 Excel 的本地绝对路径
INPUT_EXCEL_PATH = r"D:\Desktop\Testcase_YoyoTopics_104Agents_with_Chinese_requests.xlsx"
#==========================================
# 4. 性能与并发配置 (Performance & Concurrency)
# ==========================================
# 同时开启的浏览器/网页最大数量 (根据电脑性能调整，推荐 3-5)
MAX_WORKERS = 3