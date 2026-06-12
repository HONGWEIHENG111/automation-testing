import os
import time
from datetime import datetime
from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import PyPDF2
import base64
import openpyxl
import keyboard
import threading
import re
from selenium.webdriver.chrome.options import Options
from tools.state_manager import execute_state, handle_post_send
from tools.llm_evaluator import evaluate_with_deepseek
from tools.summary_generator import generate_summary_csv
from tools.agent_mapping import AGENT_NAME_MAPPING_TC, TC_TO_ENG_MAPPING
import difflib
from config import USERNAME, PASSWORD, HOME_URL, INPUT_EXCEL_PATH
# 动态生成所有合法的 Agent 名称全局列表（包含所有中英文标准名称）
ALL_VALID_AGENTS = list(AGENT_NAME_MAPPING_TC.keys())
for tc_list in AGENT_NAME_MAPPING_TC.values():
    ALL_VALID_AGENTS.extend(tc_list)
ALL_VALID_AGENTS = list(set(ALL_VALID_AGENTS))

STOP_SCRIPT = False

def listen_for_hotkey():
    """在后台运行的线程，专门监听终止键"""
    global STOP_SCRIPT
    keyboard.add_hotkey('ctrl+alt+h', show_authorship)
    # 阻塞等待按下 Ctrl+C
    keyboard.wait('ctrl+c')
    STOP_SCRIPT = True
    print("\n\n🛑 [紧急刹车] 侦测到 Ctrl+C 键按下！脚本将在当前步骤跳出，并释放浏览器控制权...\n")

def smart_sleep(seconds):
    """可随时被紧急刹车打断的睡眠函数"""
    for _ in range(int(seconds * 10)):
        if STOP_SCRIPT:
            break
        time.sleep(0.1)

def close_popups(driver):
    """专门用于检测并关闭页面上的干扰弹窗"""
    print("🔍 正在执行弹窗清理扫雷...")
    # 【目标 1】处理 Welcome 导览弹窗 (点击 Skip)
    try:
        skip_btn = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.ID, "gen-tour-welcome-skip"))
        )
        driver.execute_script("arguments[0].click();", skip_btn)
        print("✅ 已成功跳过 Welcome 导览弹窗")
        time.sleep(0.5)
    except Exception:
        pass

    # 【目标 2】处理 News/更新 弹窗
    try:
        dont_show_label = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.XPATH, "//label[@for='dontShowAgain']"))
        )
        driver.execute_script("arguments[0].click();", dont_show_label)
        print("✅ 已勾选 'Don't show this again'")
        time.sleep(0.5)

        close_btn = driver.find_element(By.CLASS_NAME, "credit-popup-close")
        driver.execute_script("arguments[0].click();", close_btn)
        print("✅ 已成功关闭 News 弹窗")
    except Exception:
        pass

def get_test_data_from_excel(excel_path):
    """从 Excel 中智能寻找 'Request' 和 'Selected Agent'，并提取其正下方的所有内容"""
    questions = []
    selected_agents = []
    filenames = []
    selected_languages = []

    if not os.path.exists(excel_path):
        print(f"[致命错误] 找不到 Excel 文件: {excel_path}")
        return questions, selected_agents, filenames, selected_languages

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

        request_col = None
        agent_col = None
        filename_col = None
        language_col = None
        header_row = None

        # 1. 遍历表格，寻找表头
        for row in ws.iter_rows():
            for cell in row:
                if cell.value:
                    val = str(cell.value).strip().lower()
                    if val == "request":
                        header_row = cell.row
                        request_col = cell.column
                    elif val == "selected agent":
                        header_row = cell.row  # 假设它们在同一行
                        agent_col = cell.column
                    elif val in ["filename", "file name", "file_name"]:  # 兼容不同写法
                        header_row = cell.row
                        filename_col = cell.column
                    else:
                        # 标准和常见变体词库
                        lang_keywords = ["selected language", "selected_language", "language", "lang", "目标语言",
                                         "测试语言", "ui语言", "ui language"]
                        is_language_column = False

                        # 1. 精确词库匹配
                        if val in lang_keywords:
                            is_language_column = True

                        # 2. 子串包含匹配（只要包含核心词根就命中）
                        elif "language" in val or "lang" in val or "语言" in val:
                            is_language_column = True

                        # 3. 模糊相似度匹配（容忍测试人员手误拼错，如 "langauge"）
                        else:
                            matches = difflib.get_close_matches(val, lang_keywords, n=1, cutoff=0.6)
                            if matches:
                                is_language_column = True

                        # 如果通过以上任一防线确认是语言列，则记录列号
                        if is_language_column:
                            header_row = cell.row
                            language_col = cell.column
            # 只要发现了 Request 或 Filename 任意一个核心锚点，就确认是表头并跳出扫描。
            # 这样既不会被可能在第一行出现的随机备注（如 "默认Language"）骗到，也能完整扫描完这一行的所有列。
            if request_col is not None or filename_col is not None:
                break

        if not request_col and not filename_col:
            print(f"[致命错误] 在 Excel({excel_path}) 中找不到 'Request' 也找不到 'Filename' 表头！无法提供任何测试输入。")
            return questions, selected_agents, filenames, selected_languages

        if not filename_col:
            print(f"[降级提示] 找不到 'Filename' 表头，将自动降级为【纯文本模式】运行。")

        if not agent_col:
            print(f"[降级提示] 找不到 'Selected Agent' 表头，将自动降级为【状态 2 (通用自动模式)】。")

        if not language_col:
            print(f"[降级提示] 找不到 'Language' 表头，将自动使用【英文 (默认)】。")

        for row in range(header_row + 1, ws.max_row + 1):
            req_val = ws.cell(row=row, column=request_col).value if request_col else ""
            file_val = ws.cell(row=row, column=filename_col).value if filename_col else ""

            req_str = str(req_val).strip() if req_val else ""
            file_str = str(file_val).strip() if file_val else ""

            # 只要 Request 不为空 或者 Filename 不为空，就认为这是一条有效数据
            if req_str or file_str:
                questions.append(req_str)
                filenames.append(file_str)
                # 读取对应的 Agent
                if agent_col:
                    agent_val = ws.cell(row=row, column=agent_col).value
                    selected_agents.append(str(agent_val).strip() if agent_val else "")
                else:
                    selected_agents.append("")

                # 读取对应的 Language
                if language_col:
                    lang_val = ws.cell(row=row, column=language_col).value
                    selected_languages.append(str(lang_val).strip() if lang_val else "")
                else:
                    selected_languages.append("")
    except Exception as e:
        print(f"[格式错误] 读取 Excel 文件失败: {e}")

    return questions, selected_agents, filenames, selected_languages

def show_authorship():
    secret_msg = base64.b64decode(b'Q29kZSBBdXRob3I6IEhlbnJ5IEhPTkc=').decode('utf-8')
    print(f"\n\n====================================")
    print(f"🌟 {secret_msg} 🌟")
    print(f"====================================\n\n")
def read_file_content(file_path):
    """根据文件后缀读取不同格式的文件内容"""
    ext = os.path.splitext(file_path)[1].lower()
    content = ""
    try:
        # 处理 TXT 和 CSV
        if ext in ['.txt', '.csv']:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                # 兼容 Windows 常见的 GBK 编码
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()

        # 处理 DOCX
        elif ext == '.docx':
            doc = Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

        # 处理 PDF
        elif ext == '.pdf':
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                content = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])

    except Exception as e:
        print(f"❌ 读取文件 {os.path.basename(file_path)} 内容时失败: {e}")

    return content

def run_automation():
    global STOP_SCRIPT
    # 开启后台键盘监听线程
    listener_thread = threading.Thread(target=listen_for_hotkey, daemon=True)
    listener_thread.start()
    # 在这里手动修改你要运行的状态：1, 2, 3...
    CURRENT_STATE = "4"
    # ================= 动态路径配置（替代写死的绝对路径） =================
    project_dir = os.path.dirname(os.path.abspath(__file__))
    test_dir = os.path.join(project_dir, "test")

    if not os.path.exists(test_dir):
        print(f"[致命错误] 找不到测试文件夹: {test_dir}")
        return

    # 获取所有问题
    questions, selected_agents, target_filenames, selected_languages = get_test_data_from_excel(INPUT_EXCEL_PATH)
    if not questions:
        print("未提取到任何问题，程序终止。")
        return
    # 获取 test 文件夹中支持的文件，并进行【按文件名排序】
        # 严格按照网站报错提示的格式进行过滤
    valid_extensions = ('.pdf', '.docx', '.csv', '.txt')
    test_files = [f for f in os.listdir(test_dir) if f.lower().endswith(valid_extensions)]

    # 排序以保证和 Word 里面的顺序完全一致
    test_files.sort()

    if len(test_files) != len(questions):
        print(f"文件夹中的文件数量 ({len(test_files)}) 与 Excel 中的问题数量 ({len(questions)}) 不匹配！")

    # 启动浏览器
    try:
        print("🔍 尝试启动首选浏览器：Google Chrome...")
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument('--ignore-certificate-errors') # 👈 添加这一行即可无视警告
        # 禁止浏览器后台休眠和节流
        chrome_options.add_argument('--disable-background-timer-throttling')
        chrome_options.add_argument('--disable-backgrounding-occluded-windows')
        chrome_options.add_argument('--disable-renderer-backgrounding')
        # chrome_options.add_argument('--headless=new') #当完全没有问题之后可以使用这个无头模式，不用盯着界面看，节省70%内存，可以调高worknum
        # chrome_options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        print("✅ 成功启动 Google Chrome！")
    except Exception as e_chrome:
        print(f"⚠️ Chrome 启动失败: {e_chrome}")
        print("🔄 正在尝试启动备用浏览器：Microsoft Edge...")
        try:
            edge_options = EdgeOptions()
            edge_options.add_experimental_option("detach", True)
            edge_options.add_argument('--ignore-certificate-errors')
            driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)
            print("✅ 成功启动 Microsoft Edge！")
        except Exception as e_edge:
            print(f"❌ [致命错误] Chrome 和 Edge 均启动失败，请检查浏览器安装环境！\n错误信息: {e_edge}")
            return
    driver.get(HOME_URL)
    time.sleep(1)  # 等待页面初始加载

    # ================= 自动登录逻辑 =================
    print("🔑 正在执行自动登录...")
    try:
        # 1. 点击导航栏的 Login (此处可以使用原生 click，因为刚启动一般还在前台，但稳妥起见可以用显式等待)
        login_nav = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "nav-login"))
        )
        driver.execute_script("arguments[0].click();", login_nav)

        # 2. 等待账号密码框出现，并使用 JS 注入
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "sso-username"))
        )
        password_input = driver.find_element(By.ID, "sso-password")

        # JS 注入账号密码
        driver.execute_script("arguments[0].value = arguments[1];", username_input, USERNAME)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", username_input)

        driver.execute_script("arguments[0].value = arguments[1];", password_input, PASSWORD)
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", password_input)

        # 3. JS 点击提交按钮
        submit_btn = driver.find_element(By.CLASS_NAME, "sso-submit-btn")
        driver.execute_script("arguments[0].click();", submit_btn)
        close_popups(driver)
        print("✅ 登录动作已提交，等待页面刷新...")
        time.sleep(1)  # 等待登录成功后的页面跳转和加载
        close_popups(driver)
    except Exception as e:
        print(f"❌ [登录失败] 请检查页面是否卡顿或元素发生变化: {e}")
        driver.quit()
        return

    # ================= 初始化 Excel 结果文件 =================
    # 1. 提取原始测试用例的文件名（例如把 "Testcase_20260520_1.xlsx" 变成 "Testcase_20260520_1"）
    base_testcase_name = os.path.splitext(os.path.basename(INPUT_EXCEL_PATH))[0]

    # 2. 生成当前时间戳（格式例如：20260526_090339）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 3. 动态组合最终的文件名
    dynamic_filename = f"evaluation_results_{base_testcase_name}_{timestamp}.xlsx"
    excel_dir = os.path.join(project_dir, "Evaluation_Results")
    os.makedirs(excel_dir, exist_ok=True)
    excel_path = os.path.join(excel_dir, dynamic_filename)
    if not os.path.exists(excel_path):
        wb = openpyxl.Workbook()
        # 🟢 隐蔽水印：修改 Excel 文件的底层元数据
        wb.properties.creator = "Henry HONG "
        wb.properties.description = "Authored by Henry HONG. "
        ws = wb.active
        ws.title = "Evaluation Results"
            # 写入表头
        ws.append(["label","Request","filename", "Crash", "Tester Expectation", "Selected Language","Input Language", "Output Language", "Language Overall Status", "answer", "shared link", "DeepSeek评价内容", "Selected agent", "Reference Link", "Document Contain[1][2][3]", "Preparation Time", "Completion Time", "Timeout_States"])
        wb.save(excel_path)
        print(f"📊 已创建评价结果 Excel 文件: {excel_path}")
    # ===============================================================
    '''
    情况1：有request，selected agent， filename（且test中有对应文件）可以正常运行
    情况2：有request，selected agent， filename（且test中无对应文件），转为纯文本（只看request)运行并给出提示
    情况3：有request，selected agent，无filename，转为纯文本（只看request)运行并给出提示
    情况4：有request，filename（且test中有对应文件），无selected agent，按照状态2运行
    情况5：有request，filename（且test中无对应文件），无selected agent，按照状态2转为纯文本（只看request)运行并给出提示
    情况6：有request，无filename，无selected agent，按照状态2转为纯文本（只看request)运行并给出提示
    情况7：无request，有selected agent， filename（且test中有对应文件）可以正常运行
    情况8：无request，有selected agent， filename（且test中无对应文件），报错并且跳过进行下一项
    情况9：无request，有selected agent， 无filename，报错并且跳过进行下一项
    情况10：无request，无selected agent，有filename（且test中有对应文件），按照状态2运行
    情况11：无request，无selected agent，有filename（且test中无对应文件），报错并且跳过进行下一项
    情况12：无request，无selected agent，无filename，报错并且跳过进行下一项
    '''

        # 按照 Excel 的行数顺序遍历执行
    for i in range(len(questions)):
        if STOP_SCRIPT:
            print("🛑 收到中止指令，停止处理后续文件。")
            break
        timeout_status = "No"
        # 🌟每次准备处理新题目时，先看一眼浏览器还在不在
        try:
            _ = driver.window_handles
        except Exception:
            print("\n🚨 侦测到浏览器窗口已被手动关闭，脚本终止运行！")
            break

        question_text = questions[i]
        target_agent = selected_agents[i]
        filename = target_filenames[i]
        target_language = selected_languages[i]

        actual_file_name = None
        file_path = None
        has_file_input = False

        # 1. ==== 处理文件查找与降级逻辑 ====
        if filename and str(filename).strip():
            target_base_name = str(filename).strip().lower()

            # 遍历 test 文件夹寻找匹配
            for f in os.listdir(test_dir):
                base_name_in_folder, ext = os.path.splitext(f)
                if base_name_in_folder.strip().lower() == target_base_name:
                    actual_file_name = f
                    file_path = os.path.join(test_dir, actual_file_name)
                    has_file_input = True
                    break

            # 核心逻辑：如果找不到文件
            if not has_file_input:
                if question_text and str(question_text).strip():
                    print(f"\n⚠️ [降级运行] 找不到文件 '{filename}'，但存在 Request，降级为仅发送问题的纯文本模式！")
                    # 不 continue，继续往下跑，只是 has_file_input 依然是 False
                else:
                    print(f"\n❌ [跳过任务] 找不到文件 '{filename}' 且没有 Request，无法继续，跳过此条任务！")
                    continue
        else:
            # 如果 Excel 里没写 filename，再双重确认下有没有 request
            if not question_text or not str(question_text).strip():
                print(f"\n❌ [跳过任务] Request 和 Filename 同时为空，跳过此条任务！")
                continue

        # 标志位：没找到文件，或者根本没提供 filename，那就是纯文本模式
        is_text_only = not has_file_input

        # 2. ==== 核心状态分配逻辑 ====
        # 只要指定了 Agent 就用状态 4，否则用状态 2
        if target_agent and str(target_agent).strip():
            CURRENT_STATE = "4"
        else:
            CURRENT_STATE = "2"


        target_agent_raw = str(target_agent).strip() if target_agent else ""
        # ================= 拼写纠错与智能降级逻辑 =================
        if target_agent_raw and target_agent_raw not in ALL_VALID_AGENTS:
            lower_agents_mapping = {agent.lower(): agent for agent in ALL_VALID_AGENTS}
            target_agent_lower = target_agent_raw.lower()

            # 第一步防禦：嘗試「無視大小寫的精確匹配」（解決 ceo -> CEO, data analyst -> Data Analyst）
            if target_agent_lower in lower_agents_mapping:
                correct_name = lower_agents_mapping[target_agent_lower]
                print(
                    f"[{i + 1}] 🔧 [大小寫修正]: 偵測到大小寫不標準 '{target_agent_raw}'，已自動修正為 '{correct_name}'")
                target_agent_raw = correct_name
            else:
                # 第二步防禦：如果大小寫一致也找不到，再走「無視大小寫的模糊匹配」（解決拼寫錯誤，如 data analist）
                matches = difflib.get_close_matches(target_agent_lower, list(lower_agents_mapping.keys()), n=1,
                                                    cutoff=0.6)

                if matches:
                    correct_name = lower_agents_mapping[matches[0]]
                    print(
                        f"[{i + 1}] 🔧 [智能糾錯]: 偵測到拼寫錯誤 '{target_agent_raw}'，已自動糾正為合法的 '{correct_name}'")
                    target_agent_raw = correct_name  # 用正確的名字覆蓋
                else:
                    print(
                        f"[{i + 1}] ⚠️ [降級警告]: 拼写错误离谱，无法识别 Agent '{target_agent_raw}'，自動降級為通用模式 (State 2)！")
                    target_agent_raw = ""  # 清空名字
                    CURRENT_STATE = "2"  # 強制降級為不選 Agent 的狀態
        # ===================================================================
        # 判断当前的 UI 语言是否为繁中
        tc_keywords = ["繁中", "繁体", "繁體", "traditional chinese", "tc", "zh-tw"]
        is_tc_ui = target_language and any(keyword in str(target_language).lower() for keyword in tc_keywords)
        # 升级为候选列表，支持多个备用词
        search_candidates = []
        if target_agent_raw:
            if is_tc_ui:
                # -------------------------------------------------------------
                # 【情况 A】Selected Language 是繁中
                # -------------------------------------------------------------
                if target_agent_raw in AGENT_NAME_MAPPING_TC:
                    # 逻辑 1：Agent 输入是英文 -> 获取整个繁中候选列表 (例如 ["談判策略", "談判方案"])
                    search_candidates = AGENT_NAME_MAPPING_TC[target_agent_raw]
                    print(f"   🔄 [Agent 转换]: 繁中 UI 匹配到英文输入，载入候选列表 -> {search_candidates}")
                else:
                    # 逻辑 2：Agent 输入是繁中 -> 直接用繁中搜索
                    search_candidates = [target_agent_raw]
                    print(f"   🎯 [Agent 保持]: 繁中 UI 匹配到繁中输入，直接搜索 -> '{target_agent_raw}'")
            else:
                # -------------------------------------------------------------
                # 【情况 B】Selected Language 是英文或者空白
                # -------------------------------------------------------------
                if target_agent_raw in TC_TO_ENG_MAPPING:
                # 逻辑 4：Agent 输入是繁中 -> 繁中转英文搜索
                    search_candidates = [TC_TO_ENG_MAPPING[target_agent_raw]]
                    print(f"   🔄 [Agent 转换]: 英文 UI 匹配到繁中输入，已自动转为英文 -> '{search_candidates[0]}'")
                else:
                # 逻辑 3：Agent 输入是英文 -> 直接用英文搜索
                    search_candidates = [target_agent_raw]
                    print(f"   🎯 [Agent 保持]: 英文 UI 匹配到英文输入，直接搜索 -> '{target_agent_raw}'")
        # ==== 打印信息，方便你监控 ====
        print(f"\n--- 进度: {i + 1}/{len(questions)} ---")
        print(f"❓ 输入问题: {question_text if question_text else '【无文本，仅上传文件】'}")
        print(
            f"🤖 目标Agent: {target_agent if target_agent and str(target_agent).strip() else '【未指定，自动使用状态 2】'}")
        print(f"⚙️ 最终执行状态: {CURRENT_STATE}")

        if not is_text_only:
            print(f"📄 准备上传文件: {actual_file_name}")
        else:
            print(f"📄 纯文本/降级模式，无需上传文件。")

        # --- 网页自动化交互逻辑 ---
        try:
            # 只有明确指定了繁中或繁体，才切换繁中；其他所有情况（包括为空、英文、乱码等）全都默认切英文
            if is_tc_ui:
                print("   🌐 [语言切换]: 正在切换为 繁中...")
                # 1. 点击 Language 容器展开下拉菜单
                lang_container = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "language-container"))
                )
                driver.execute_script("arguments[0].click();", lang_container)
                time.sleep(1)

                # 2. 点击 繁中
                tc_option = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//a[contains(@class, 'language-option') and text()='繁中']"))
                )
                driver.execute_script("arguments[0].click();", tc_option)
                time.sleep(1.5)  # 必须留足时间等待新页面加载

            else:
                print("   🌐 [语言切换]: 正在切换为 英文 (默认)...")
                # 1. 点击 Language 容器展开下拉菜单
                lang_container = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "language-container"))
                )
                driver.execute_script("arguments[0].click();", lang_container)
                time.sleep(1)

                # 2. 点击 ENG
                eng_option = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//a[contains(@class, 'language-option') and text()='ENG']"))
                )
                driver.execute_script("arguments[0].click();", eng_option)
                time.sleep(1)
            #  2：只有在非纯文本模式下，才执行上传文件动作
            if not is_text_only and file_path:
                file_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                )
                file_input.send_keys(file_path)
                time.sleep(1)  # 等待前端读取文件
            else:
                print("   📄 [跳过]: 纯文本模式，无需上传文件。")

            # 动作 2：输入问题文本（用显式等待DOM存在 + JS异步注入，完美兼容窗口最小化）
            try:
                # 1. 显式等待元素加载到 DOM 树中（不要求窗口必须在前台可见）
                text_area = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "background-info"))
                )

                # 2. 使用 JavaScript 强行清空原文本
                driver.execute_script("arguments[0].value = '';", text_area)

                # 3. 使用 JavaScript 强行注入新的问题文本
                driver.execute_script("arguments[0].value = arguments[1];", text_area, question_text)

                # 4. 关键核心：向输入框连续派发 input 和 change 事件，确保 Vue/React 等前端框架能抓取到新数据
                driver.execute_script("""
                    var input = arguments[0];
                    var lastValue = input.value;
                    input.value = arguments[1];
                    var event = new Event('input', { bubbles: true });
                    // 破解 React 16+ 内部的值追踪器
                    var tracker = input._valueTracker;
                    if (tracker) { tracker.setValue(lastValue); }
                    input.dispatchEvent(event);
                """, text_area, question_text)
                driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", text_area)

                print("   ✅ 已成功在后台写入问题文本。")
                time.sleep(1)

            except Exception as input_err:
                print(f"   ❌ 输入问题文本时发生异常: {input_err}")
                raise input_err

            # 动作3，按齿轮
            print("   🔍 [通用步骤]: 正在点击设置齿轮图标...")
            gear_btn = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'bi-gear-fill')]"))
            )
            driver.execute_script("arguments[0].click();", gear_btn)
            time.sleep(0.5)

            # ================= 根据不同状态执行完全独立的点击逻辑 =================
            if search_candidates:
                agent_success = False
                last_err = None
                for candidate in search_candidates:
                    try:
                        print(f"   👉 正在尝试选中 Agent: '{candidate}' ...")
                        execute_state(driver, CURRENT_STATE, candidate)
                        agent_success = True
                        print(f"   ✅ 成功找到并应用 Agent: '{candidate}'")
                        break
                    except Exception as e:
                        print(f"   ⚠️ 当前网站未找到 '{candidate}'，准备尝试下一个候选词...")
                        last_err = e

                if not agent_success:
                    print(f"   ⚠️ [保底机制触发]: 候选 Agent {search_candidates} 均无法找到！已自动降级为 State 2。")
                    CURRENT_STATE = "2"
                    target_agent = str(target_agent) + " (未找到，降级为 State 2)"
                    execute_state(driver, CURRENT_STATE)
            else:
                print("   ℹ️ 目标 Agent 为空，直接应用通用模式 (State 2)...")
                execute_state(driver, CURRENT_STATE)

        # Apply settings
            print("   🔍 [通用步骤]: 面板配置完毕，正在点击 Apply Settings...")
            apply_btn = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "global-settings-apply-btn"))
            )
            driver.execute_script("arguments[0].click();", apply_btn)
            time.sleep(1)
        # 点击发送
            print("   🚀 [通用步骤]: 正在点击发送/生成按钮...")
            submit_btn = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'btn-enter')]"))
            )
            driver.execute_script("arguments[0].click();", submit_btn)
        # ================================================================
        # 发送后的特殊操作（针对状态 1）
            handle_post_send(driver, CURRENT_STATE)
            print("   ✅ 已成功进入自动生成等待环节...")
        # ================================================================

        # 动作 4：等待生成结果
            print("⏳ 正在智能监控 AI 生成进度...")
            # 初始化当前用例的超时状态 ---
            timeout_status = "No"
            try:
                smart_sleep(5)  # 先强制等待 5 秒，让前端动画和网络请求先跑起来

                last_length = 0
                stable_count = 0
                max_wait_loops = 300  # 最大循环次数（约 300 秒，防止死循环）
                required_stable_seconds = 5  # 允许中途卡顿 5 秒！可以根据实际情况调大
                # 最多循环检测 60 次（约 300 秒超时限制）
                for _ in range(max_wait_loops):
                    if STOP_SCRIPT:
                        print("🛑 收到中止指令，立即停止网页监控！")
                        break
                    # 🌟监控途中检查浏览器状态，防止关闭后死等 300 秒
                    try:
                        _ = driver.window_handles
                    except Exception:
                        print("\n🚨 监控期间侦测到浏览器窗口被关闭，立即退出！")
                        STOP_SCRIPT = True
                        break
                    time.sleep(1)  # 每秒检查一次

                    current_length = 0
                    try:
                        # ★ 核心：不要抓整个 body，只抓当前正在生成的那个回答的字数
                        previews = driver.find_elements(By.ID, "preview")
                        for p in reversed(previews):
                            if p.is_displayed():
                                actual_text = p.get_attribute("innerText") or p.get_attribute("textContent") or ""
                                current_length = len(actual_text.strip())
                                break  # 找到了最新的回答，获取长度后跳出
                    except Exception:
                        pass  # 忽略查找元素时偶发的 DOM 刷新错误

                # 只有当当前长度大于 0（说明 AI 已经开始说话了），且和上次一样，才算作稳定
                    if current_length > 0 and current_length == last_length:
                        stable_count += 1
                    else:
                        stable_count = 0  # 只要字数变化，或者还没开始吐字，就重置计时器

                    last_length = current_length

                # 如果连续 3 秒字数完全没有变化，判定为生成结束
                    if stable_count >= required_stable_seconds:
                        print(
                            f"✅ 回答文本已连续 {required_stable_seconds} 秒无变化，判定生成彻底完成！最终字数: {current_length}")
                        break
                else:
                    if not STOP_SCRIPT:  # 确保不是因为你按了 Ctrl+C 退出的
                        print("⚠️ 警告：监控达到 300 秒上限，生成总时间超时！")
                        timeout_status = "yes (总时间超时)"
            except Exception as wait_error:
                print(f"⚠️ 智能监控发生异常，强制继续执行: {wait_error}")
                time.sleep(2)  # 保底等待
            if STOP_SCRIPT:
                print("🛑 侦测到中止指令，已跳过 DeepSeek 评价并中止后续任务！")
                break
            # ================= 提取当前页面的链接 =================
            current_page_url = driver.current_url
            print(f"🔗 已抓取当前回答shared link: {current_page_url}")
        # ================= 提取回答并发送给 DeepSeek 进行评价并保存 Excel =================
            print("📥 正在提取生成的回答并读取原文件...")
            try:
            # 1. 提取网页回答 (增强版：延长等待，强制倒序查找最新生成的 preview)
                print("   ⏳ 正在等待并提取最新的可见回答 (最多等待 60 秒)...")

                def get_valid_preview(d):
                    # 🌟防止浏览器关闭后，WebDriverWait 在这里硬挺 60 秒
                    try:
                        _ = d.window_handles
                    except Exception:
                        return "BROWSER_CLOSED"
                    previews = d.find_elements(By.ID, "preview")

                # 💡 关键修改：用 reversed() 倒序遍历！
                # 因为如果不刷新网页，页面上会堆积 3 个 preview。
                # 最新的回答一定是在网页的最底层（列表的最后面），倒着找能 100% 避开前两个旧回答的干扰。
                    for p in reversed(previews):
                        actual_content = p.get_attribute("innerText") or p.get_attribute("textContent") or ""
                        if len(actual_content.strip()) > 0:
                            return actual_content
                    return False

                try:
                    answer_text = WebDriverWait(driver, 60).until(get_valid_preview)
                    if answer_text == "BROWSER_CLOSED":
                        print("🚨 提取内容时侦测到浏览器已关闭，终止后续任务！")
                        break
                except Exception:
                    answer_text = ""

            # 保底措施：如果 60 秒到了实在没抓到
                if not answer_text or not answer_text.strip():
                    print("   ⚠️ 警告：等待 60 秒后依然未抓取到内容！")
                # 加一个额外的调试信息，看看页面上到底有几个 preview
                    debug_previews = driver.find_elements(By.ID, "preview")
                    print(f"   🔍 调试信息：当前页面共有 {len(debug_previews)} 个 ID 为 preview 的元素。")

                    answer_text = "提取文本失败/为空"
                    if timeout_status == "No":
                        timeout_status = "yes (生成超时)"
                    else:
                        timeout_status += " & yes (生成超时)"
                else:
                    print(f"   ✅ 成功提取到回答，长度: {len(answer_text)} 字符")

                prep_time = "N/A"
                comp_time = "N/A"
                try:
                    # 1. 强制等待 3 秒，等组件完全挂载到 React 树上
                    time.sleep(3)

                    # 2. 注入核弹级 JS：全方位透视 DOM 与 React 底层内存树
                    times_info = driver.execute_script("""
                                    try {
                                        var prep = 'N/A', comp = 'N/A';
                                        var debugLog = [];
                                        // 锁定我们要找的滚动组件标签
                                        var flows = document.querySelectorAll('number-flow-react, number-flow');
                                        debugLog.push('找到标签数: ' + flows.length);
                                        var vals = [];
    
                                        for(var i=0; i<flows.length; i++) {
                                            var el = flows[i];
                                            var val = null;
    
                                            // 第一层：尝试常规 DOM 属性
                                            if (el.value !== undefined) val = el.value;
                                            else if (el.getAttribute('value')) val = el.getAttribute('value');
                                            else if (el.getAttribute('aria-valuenow')) val = el.getAttribute('aria-valuenow');
    
                                            // 第二层：入侵 React 16-18+ 的 Fiber 内存树，向上回溯寻找 memoizedProps
                                            if (val == null) {
                                                var fiberKey = Object.keys(el).find(k => k.startsWith('__reactFiber$'));
                                                if (fiberKey) {
                                                    var curr = el[fiberKey];
                                                    // 向上回溯 5 层，寻找包含 value 变量的节点
                                                    for (var j=0; j<5; j++) {
                                                        if (curr && curr.memoizedProps && curr.memoizedProps.value !== undefined) {
                                                            val = curr.memoizedProps.value;
                                                            debugLog.push('通过Fiber提取成功');
                                                            break;
                                                        }
                                                        if (curr) curr = curr.return;
                                                    }
                                                }
                                            }
    
                                            // 第三层：入侵 React 17+ 的 Props 内存
                                            if (val == null) {
                                                var propsKey = Object.keys(el).find(k => k.startsWith('__reactProps$'));
                                                if (propsKey && el[propsKey] && el[propsKey].value !== undefined) {
                                                    val = el[propsKey].value;
                                                    debugLog.push('通过Props提取成功');
                                                }
                                            }
    
                                            if (val !== null) vals.push(val);
                                        }
    
                                        if (vals.length >= 1) prep = vals[0] + 's';
                                        if (vals.length >= 2) comp = vals[1] + 's';
    
                                        return { prep: prep, comp: comp, debug: debugLog.join(' | ') };
                                    } catch(err) {
                                        return { prep: 'N/A', comp: 'N/A', debug: 'JS报错: ' + err.message };
                                    }
                                """)

                    if times_info:
                        prep_time = times_info.get("prep", "N/A")
                        comp_time = times_info.get("comp", "N/A")
                        debug_info = times_info.get("debug", "")

                        # 第四层极限保底：如果连 React 内存里都没有，强行用正则刮取底层 HTML 代码
                        if prep_time == "N/A":
                            html = driver.execute_script("return document.body.innerHTML;")
                            # 匹配类似 value="99.5" 的隐藏属性
                            html_matches = re.findall(r'<number-flow[^>]*?(?:value|aria-valuenow)=["\']?([0-9.]+)["\']?',
                                                      html, re.IGNORECASE)
                            if len(html_matches) >= 1:
                                prep_time = html_matches[0] + "s"
                                debug_info += " | 触发底层HTML正则兜底"
                            if len(html_matches) >= 2:
                                comp_time = html_matches[1] + "s"

                        print(
                            f"[{i + 1}] ⏱️ 提取时间 -> 准备耗时: {prep_time}, 完成耗时: {comp_time} (底层诊断: {debug_info})")
                    else:
                        print(f"[{i + 1}] ⏱️ 提取失败：JS 脚本未返回任何数据。")

                except Exception as time_err:
                    print(f"[{i + 1}] ⚠️ 提取时间发生代码异常: {time_err}")

            # 2. 提取本地文档内容
                if is_text_only:
                    file_content = "【无原始文档，用户仅提供了纯文本提问，请仅根据问题本身评估回答是否准确且符合逻辑】"
                else:
                    file_content = read_file_content(file_path)

                if "yes" in timeout_status.lower():
                    print("   ⚠️ 侦测到超时，跳过 DeepSeek 评价，直接记录为超时失败...")
                    eval_results = {
                        "tester_expectation": "Failed",
                        "input_language": "N/A",
                        "output_language": "N/A",
                        "language_status": "Failed",
                        "evaluation": "生成超时或提取失败，未能获取有效回答。",
                        "reference_link": "N/A",
                        "document_contain_citations": "None"
                    }
                else:
                    print("🤖 成功提取网页回答，正在等待 DeepSeek 进行语言检测与质量评价...")
                    eval_results = evaluate_with_deepseek(question_text, file_content, answer_text, target_language)
                tester_exp = eval_results.get("tester_expectation", "Unknown")
                input_lang = eval_results.get("input_language", "Unknown")
                output_lang = eval_results.get("output_language", "Unknown")
                lang_status = eval_results.get("language_status", "Unknown")
                evaluation_text = eval_results.get("evaluation", "Unknown")  # 移除了 raw_result，统一用 "Unknown" 作为后备
                ref_link = eval_results.get("reference_link", "N/A")
                doc_contain = eval_results.get("document_contain_citations", "None")
            # ===============================================================

            # 5. 将结果追加写入到 Excel 文件中
                try:
                    wb = openpyxl.load_workbook(excel_path)
                    ws = wb.active
                #截断过长的回答，只保留前 100 个字符】
                    short_answer = answer_text[:100] + "..." if len(answer_text) > 100 else answer_text
                # 🌟 确保 filename 为空时填入空字符串，而不是报错或填入 None
                    display_filename = filename if filename else ""
                    display_target_language = target_language if target_language else "N/A"
                # 写入一行数据：问题, 文件名, 回答, 评价, 使用的模型
                # 【修改这里：按照表头顺序，加入解析出的语言信息】
                    ws.append([
                        i + 1,  # 填入 label 序号 (1, 2, 3...)
                        question_text,  # 原Request内容
                        display_filename,  # 测试filename
                        "No",  # Crash (第4列，正常情况写 No)
                        tester_exp,  # Tester Expectation
                        display_target_language,
                        input_lang,  # Input Language
                        output_lang,  # Output Language
                        lang_status,  # Language Overall Status
                        short_answer,  # 生成的回答
                        current_page_url,  # 保存刚才抓取的网页链接】
                        evaluation_text,  # DeepSeek评价内容
                        target_agent if target_agent else "未指定",  # agent
                        ref_link,  # Reference Link
                        doc_contain,  # Document Contain[1][2][3]
                        prep_time,  # Preparation Time
                        comp_time,  # Completion Time
                        timeout_status, # 超时状态
                    ])
                    wb.save(excel_path)
                    print("✅ 评价结果已成功写入 Excel 文件。")
                except Exception as excel_err:
                    print(f"❌ 写入 Excel 时失败: {excel_err}")

            except Exception as extract_error:
                print(f"❌ 提取文本、评价或保存数据时发生错误: {extract_error}")

            # ================= 每次处理完后强制重置为英文 (新增) =================
            try:
                print("   🌐 [语言重置]: 正在准备返回前重置为英文...")
                    # 1. 查找并点击地球仪语言图标
                globe_icon = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "language-icon"))
                )
                driver.execute_script("arguments[0].click();", globe_icon)
                time.sleep(1)  # 等待下拉菜单展开动画

                # 2. 查找并点击 ENG 选项
                eng_reset_option = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//a[contains(@class, 'language-option') and text()='ENG']"))
                )
                driver.execute_script("arguments[0].click();", eng_reset_option)
                print("   ✅ 成功点击 ENG 重置语言")
            except Exception as reset_lang_err:
                # 如果找不到图标或点击失败，不中断主流程，直接打印警告并继续执行后续的返回首页动作
                print(f"   ⚠️ 重置语言步骤失败 (可能页面已改变或元素不可见)，将继续执行返回首页: {reset_lang_err}")
            # ==================================================================
        # ================= 点击主页 Logo 返回 =================
            try:
                print("   🏠 正在尝试通过点击 Logo 返回首页...")
                #将等待时间缩短为 7 秒，找不到就赶紧刷新，不浪费时间
                home_btn = WebDriverWait(driver, 7).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "image-home"))
                )
                #使用 JS 强制点击，防止被网页的其他弹窗或动画层遮挡
                driver.execute_script("arguments[0].click();", home_btn)
                print("   ✅ 成功通过点击 Logo 返回")
            except Exception:
                # 当找不到 image-home 元素，或者点击失败时，直接请求网址
                print("   ⚠️ 未识别到主页 Logo，正在清理会话并重新请求首页...")
                try:
                    driver.execute_script("window.sessionStorage.clear();")
                except:
                    pass
                driver.get(HOME_URL)

                # 留出一点时间让页面完全加载，准备迎接下一个文件
            time.sleep(1)
            close_popups(driver)
            if (i + 1) % 50 == 0:
                print(f"\n[{i + 1}] 🔄 触发内存保护：已连续运行 50 次，正在重启浏览器释放内存...")
                try:
                    driver.quit()
                except:
                    pass

                # 1. 重新拉起干净的浏览器实例
                driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
                driver.get(HOME_URL)
                time.sleep(2)

                # 2. 重新执行全自动登录逻辑
                print(f"[{i + 1}] 🔑 正在为新浏览器重新执行登录...")
                try:
                    login_nav = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "nav-login"))
                    )
                    driver.execute_script("arguments[0].click();", login_nav)

                    username_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "sso-username"))
                    )
                    password_input = driver.find_element(By.ID, "sso-password")

                    driver.execute_script("arguments[0].value = arguments[1];", username_input, USERNAME)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                                          username_input)
                    driver.execute_script("arguments[0].value = arguments[1];", password_input, PASSWORD)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                                          password_input)

                    submit_btn = driver.find_element(By.CLASS_NAME, "sso-submit-btn")
                    driver.execute_script("arguments[0].click();", submit_btn)

                    close_popups(driver)
                    print(f"[{i + 1}] ✅ 内存释放完毕，新浏览器登录成功，继续后续任务...")
                    time.sleep(1)
                    close_popups(driver)
                except Exception as re_login_err:
                    print(f"[{i + 1}] ❌ 重启浏览器后自动登录失败: {re_login_err}，终止后续任务。")
                    break
        # ==================================================

        except Exception as e:
            print(f"❌ [报错] 处理 {filename if filename else '纯文本'} 时发生错误: {e}")
            error_msg = str(e).lower()
            if "window already closed" in error_msg or "target window already closed" in error_msg or "no such window" in error_msg:
                print("🚨 侦测到浏览器窗口已被手动关闭，脚本放弃后续所有任务！")
                break
                # ================= 💡 新增：兜底写入 Excel 逻辑 =================
            print("   📝 正在将崩溃记录写入 Excel，防止数据漏行...")
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active

                # 容错处理：确保空值转为空字符串
                display_filename = filename if filename else ""
                display_target_language = target_language if target_language else "N/A"
                safe_agent = target_agent if target_agent else "未指定"
                crash_reason = f"Automation Crash: {str(e)[:150]}"
                # 截断报错信息，防止过长撑爆 Excel 单元格
                error_info = f"⚠️ 自动化执行崩溃: {str(e)[:150]}"

                # 严格按照表头顺序，追加一行失败的“尸体”记录
                ws.append([
                        i + 1,  # label 序号
                        question_text,  # Request
                        display_filename,  # filename
                        crash_reason,  # Crash (第4列，写入具体的报错原因)
                        "Failed",  # Tester Expectation (强行判定为失败)
                        display_target_language,  # Selected Language
                        "N/A",  # Input Language
                        "N/A",  # Output Language
                        "Failed",  # Language Overall Status
                        "【执行崩溃，未能生成回答】",  # answer
                        "N/A",  # shared link
                        error_info,  # DeepSeek评价内容 (巧妙利用这里记录报错原因)
                        safe_agent,  # Selected agent
                        "N/A",  # Reference Link
                        "None",  # Document Contain[1][2][3]
                        "N/A",  # Preparation Time
                        "N/A",  # Completion Time
                        "Crash (Error)"  # Timeout_States (标记为系统崩溃)
                ])
                wb.save(excel_path)
                print("   ✅ 崩溃记录已成功保存，数据未流失。")
            except Exception as excel_err:
                print(f"   ❌ 紧急写入兜底记录至 Excel 时失败: {excel_err}")
                # ==============================================================
            # 记录完之后，执行现场清理，重新请求首页网址，准备迎接下一个循环
            print("   🔄 正在清理现场并重新加载首页...")
            try:
                driver.get(HOME_URL)
                time.sleep(2)
            except Exception:
                pass  # 如果连刷新都报错（比如浏览器死了），就让他进入下一轮被捕捉

    print("📊 正在生成最终的 Summary 报告...")
    try:
        dynamic_csv_name = f"Summary_{base_testcase_name}_{timestamp}.csv"
        summary_dir = os.path.join(project_dir, "Summaries")  # 你可以在这里自定义 CSV 文件夹名称
        os.makedirs(summary_dir, exist_ok=True)
        output_csv = os.path.join(summary_dir, dynamic_csv_name)
        generate_summary_csv(excel_path, output_csv)
        print(f"✅ 汇总报告已生成: {output_csv}")
    except Exception as e:
        print(f"⚠️ 生成汇总报告时发生异常: {e}")
    if STOP_SCRIPT:
        print("\n🛑 任务已被手动中断！")
    else:
        print("\n✅ 文件夹内所有测试用例已全部运行完毕！")

    # 阻止控制台瞬间关闭，彻底释放并保留浏览器现场
    print("👉 自动化控制权已释放，浏览器将保持开启状态。")
    #input("按 Ctrl+C 键退出当前控制台窗口...")
    while True:
        if STOP_SCRIPT:
            print("👋 侦测到 Ctrl+C 退出指令，控制台自动关闭！")
            keyboard.unhook_all()
            break
        try:
            # 尝试获取窗口句柄，如果获取失败说明浏览器已经被你手动关闭了
            _ = driver.window_handles
            time.sleep(1)  # 每秒检查一次，不占用 CPU
        except Exception:
            print("\n👋 侦测到浏览器已被手动关闭，控制台自动退出！")
            break

if __name__ == "__main__":
    run_automation()
