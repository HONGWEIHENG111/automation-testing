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

STOP_SCRIPT = False


def listen_for_hotkey():
    """在后台运行的线程，专门监听空格键"""
    global STOP_SCRIPT
    keyboard.add_hotkey('ctrl+alt+h', show_authorship)
    # 阻塞等待按下ESC
    keyboard.wait('esc')
    STOP_SCRIPT = True
    print("\n\n🛑 [紧急刹车] 侦测到ESC键按下！脚本将在当前步骤跳出，并释放浏览器控制权...\n")

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
                    elif val in ["selected language", "selected_language"]:  # 识别 Language 表头
                        header_row = cell.row
                        language_col = cell.column
            # 如果两个表头都找到了，就跳出循环
            if request_col and agent_col:
                break

        if not request_col:
            print(f"[致命错误] 在 Excel({excel_path}) 中找不到 'Request' 表头！")
            return questions, selected_agents, filenames, selected_languages

        if not filename_col:
            print(f"[致命错误] 在 Excel({excel_path}) 中找不到 'filename' 表头！无法进行文件匹配。")
            return questions, selected_agents, filenames, selected_languages

        if not agent_col:
            print(f"[警告] 找不到 'Selected Agent' 表头，Agent 列表将返回空值。")

        for row in range(header_row + 1, ws.max_row + 1):
            req_val = ws.cell(row=row, column=request_col).value if request_col else ""
            file_val = ws.cell(row=row, column=filename_col).value if filename_col else ""

            req_str = str(req_val).strip() if req_val else ""
            file_str = str(file_val).strip() if file_val else ""

            # 只要 Request 不为空 或者 Filename 不为空，就认为这是一条有效数据
            if req_str or file_str:
                questions.append(req_str)

                # 读取对应的 Agent
                if agent_col:
                    agent_val = ws.cell(row=row, column=agent_col).value
                    selected_agents.append(str(agent_val).strip() if agent_val else "")
                else:
                    selected_agents.append("")

                # 记录 Filename
                filenames.append(file_str)
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

def run_automation(home_url, env_name):
    global STOP_SCRIPT
    # 开启后台键盘监听线程
    listener_thread = threading.Thread(target=listen_for_hotkey, daemon=True)
    listener_thread.start()


    # 在这里手动修改你要运行的状态：1, 2, 3...
    CURRENT_STATE = "4"
    # ================= 动态路径配置（替代写死的绝对路径） =================
    # 自动获取当前运行的 Python 脚本所在的文件夹路径
    USERNAME = "HenryHONG"  # 新增：替换为你的实际账号
    PASSWORD = "12345678"  # 新增：替换为你的实际密码
    project_dir = os.path.dirname(os.path.abspath(__file__))

    # 自动拼接 test 文件夹和 question.docx 的路径
    test_dir = os.path.join(project_dir, "test")
    input_excel_path = r"E:\PycharmProjects\script\Testcase_20260520_1.xlsx"
    # ===================================================================

    if not os.path.exists(test_dir):
        print(f"[致命错误] 找不到测试文件夹: {test_dir}")
        return

    # 获取所有问题
    questions, selected_agents, target_filenames, selected_languages = get_test_data_from_excel(input_excel_path)
    if not questions:
        print("未提取到任何问题，程序终止。")
        return
    if len(questions) != len(selected_agents):
        print(
            f"[致命错误] 提取到的问题数量({len(questions)})与 Agent 数量({len(selected_agents)})不一致！请检查 Excel 格式。")
        return
    # 获取 test 文件夹中支持的文件，并进行【按文件名排序】
        # 严格按照网站报错提示的格式进行过滤
    valid_extensions = ('.pdf', '.docx', '.csv', '.txt')
    test_files = [f for f in os.listdir(test_dir) if f.lower().endswith(valid_extensions)]

    # 排序以保证和 Word 里面的顺序完全一致
    test_files.sort()

    if len(test_files) != len(questions):
        print(f"[警告] 文件夹中的文件数量 ({len(test_files)}) 与 Excel 中的问题数量 ({len(questions)}) 不匹配！")

    # 启动浏览器
    try:
        print(f"[{env_name}] 🔍 尝试启动首选浏览器：Google Chrome...")
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        # 注意：顺手帮你补上了 options=chrome_options，否则你原来的 detach 不会生效
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        print("✅ 成功启动 Google Chrome！")
    except Exception as e_chrome:
        print(f"⚠️ Chrome 启动失败: {e_chrome}")
        print("🔄 正在尝试启动备用浏览器：Microsoft Edge...")
        try:
            edge_options = EdgeOptions()
            edge_options.add_experimental_option("detach", True)
            driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)
            print("✅ 成功启动 Microsoft Edge！")
        except Exception as e_edge:
            print(f"❌ [致命错误] Chrome 和 Edge 均启动失败，请检查浏览器安装环境！\n错误信息: {e_edge}")
            return
    #HOME_URL = "https://customs-demo.poffices.ai/"
    driver.get(home_url)
    time.sleep(1)  # 等待页面初始加载
    # ================= 新自动登录逻辑 =================
    print("🔑 正在执行自动登录...")
    try:
        # 1. 点击导航栏的 Login 唤出弹窗
        driver.find_element(By.ID, "nav-login").click()
        time.sleep(1)  # 等待弹窗动画加载

        # 2. 输入账号和密码
        driver.find_element(By.ID, "sso-username").send_keys(USERNAME)
        driver.find_element(By.ID, "sso-password").send_keys(PASSWORD)

        # 3. 点击提交登录按钮
        driver.find_element(By.CLASS_NAME, "sso-submit-btn").click()
        time.sleep(1)  # 等待登录页面跳转
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
    base_testcase_name = os.path.splitext(os.path.basename(input_excel_path))[0]

    # 2. 生成当前时间戳（格式例如：20260526_090339）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 3. 动态组合最终的文件名
    dynamic_filename = f"evaluation_results_{base_testcase_name}_{timestamp}.xlsx"
    excel_dir = os.path.join(project_dir, f"Evaluation_Results_{env_name}")
    os.makedirs(excel_dir, exist_ok=True)
    excel_path = os.path.join(excel_dir, dynamic_filename)
    if not os.path.exists(excel_path):
        wb = openpyxl.Workbook()
        # 🟢 隐蔽水印：修改 Excel 文件的底层元数据
        wb.properties.creator = "Henry HONG (洪伟恒)"
        wb.properties.description = "Authored by HONGWEIHENG. Tel: 17722596827"
        ws = wb.active
        ws.title = "Evaluation Results"
            # 写入表头
        ws.append(["label","Request","Tester Expectation", "filename", "Selected Language","Input Language", "Output Language", "Language Overall Status", "answer", "shared link", "DeepSeek评价内容", "Selected agent", "Reference Link", "Document Contain[1][2][3]", "Preparation Time", "Completion Time","timeout_States" ])
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
    情况12：情况11：无request，无selected agent，无filename，报错并且跳过进行下一项
    '''
        # 按照 Excel 的行数顺序遍历执行
        # ================= 全局默认状态 =================
    DEFAULT_STATE = "4"  # 如果有文件，默认使用的状态。你可以随时在这里修改
        # ================================================
        # 按照 Excel 的行数顺序遍历执行
    for i in range(len(questions)):
        if STOP_SCRIPT:
            print("🛑 收到中止指令，停止处理后续文件。")
            break
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

        # 判断当前的 UI 语言是否为繁中
        is_tc_ui = target_language and ("繁中" in str(target_language) or "繁体" in str(target_language))
        # 升级为候选列表，支持多个备用词
        search_candidates = []
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
        print(f"\n--- [{env_name}] 进度: {i + 1}/{len(questions)} ---")
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
            if target_language and ("繁中" in str(target_language) or "繁体" in str(target_language)):
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
                file_input = driver.find_element(By.XPATH, "//input[@type='file']")
                file_input.send_keys(file_path)
                time.sleep(1)  # 等待前端读取文件
            else:
                print("   📄 [跳过]: 纯文本模式，无需上传文件。")

            # 动作 2：输入问题文本
            text_area = driver.find_element(By.ID, "background-info")
            text_area.clear()
            text_area.send_keys(question_text)
            time.sleep(1)

            # 动作3，按齿轮
            print("   🔍 [通用步骤]: 正在点击设置齿轮图标...")
            gear_btn = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'bi-gear-fill')]"))
            )
            driver.execute_script("arguments[0].click();", gear_btn)
            time.sleep(0.5)

            # ================= 根据不同状态执行完全独立的点击逻辑 =================
            agent_success = False
            last_err = None

            # 【核心改动】遍历候选列表，搜不到就自动搜下一个
            for candidate in search_candidates:
                try:
                    print(f"   👉 正在尝试选中 Agent: '{candidate}' ...")
                    execute_state(driver, CURRENT_STATE, candidate)
                    agent_success = True
                    print(f"   ✅ 成功找到并应用 Agent: '{candidate}'")
                    break  # 只要成功选中一个，立刻跳出循环，继续后续流程
                except Exception as e:
                    print(f"   ⚠️ 当前网站未找到 '{candidate}'，准备尝试下一个候选词...")
                    last_err = e

            # 如果列表里的词全都试过了还是不行，抛出异常，让外层捕获并直接跳过当前用例
            if not agent_success:
                raise Exception(f"所有候选 Agent {search_candidates} 均无法找到！报错信息: {last_err}")

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
            timeout_status = "No"
            try:
                time.sleep(5)  # 先强制等待 5 秒，让前端动画和网络请求先跑起来

                last_length = 0
                stable_count = 0
                max_wait_loops = 600  # 最大循环次数（约 400 秒，防止死循环）
                required_stable_seconds = 5  # 允许中途卡顿 10 秒！可以根据实际情况调大
                # 最多循环检测 100 次（约 200 秒超时限制）
                for _ in range(max_wait_loops):
                    if STOP_SCRIPT:
                        print("🛑 收到中止指令，立即停止网页监控！")
                        break
                    # 🌟监控途中检查浏览器状态，防止关闭后死等 400 秒
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
                                current_length = len(p.text.strip())
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
                else:  # <--- 新增：注意与上方的 for 对齐
                    if not STOP_SCRIPT:
                        print("⚠️ 警告：监控达到 600 秒上限，生成总时间超时！")
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
                print("   ⏳ 正在等待并提取最新的可见回答 (最多等待 100 秒)...")

                def get_valid_preview(d):
                    # 🌟防止浏览器关闭后，WebDriverWait 在这里硬挺 80 秒
                    try:
                        _ = d.window_handles
                    except Exception:
                        return "BROWSER_CLOSED"
                    previews = d.find_elements(By.ID, "preview")

                # 💡 关键修改：用 reversed() 倒序遍历！
                # 因为如果不刷新网页，页面上会堆积 3 个 preview。
                # 最新的回答一定是在网页的最底层（列表的最后面），倒着找能 100% 避开前两个旧回答的干扰。
                    for p in reversed(previews):
                        if p.is_displayed() and len(p.text.strip()) > 0:
                            return p.text
                    return False

                try:
                    answer_text = WebDriverWait(driver, 120).until(get_valid_preview)
                    if answer_text == "BROWSER_CLOSED":
                        print("🚨 提取内容时侦测到浏览器已关闭，终止后续任务！")
                        break
                except Exception:
                    answer_text = ""

            # 保底措施：如果 100 秒到了实在没抓到
                if not answer_text or not answer_text.strip():
                    print("   ⚠️ 警告：等待 100 秒后依然未抓取到内容！")
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
                # 获取整个页面的纯文本（无视复杂的 HTML 标签嵌套）
                    time.sleep(1)
                    page_text = driver.find_element(By.TAG_NAME, "body").text

                # 用正则精准捕获数字和单位 (例如 "1s", "182.8s")
                    prep_match = re.search(r'(?:Time of preparation|準備時長:)[^\d]*([\d\.]+s?)', page_text, re.IGNORECASE)
                    if prep_match:
                        prep_time = prep_match.group(1).replace(" ", "")

                    comp_match = re.search(r'(?:Time of completion|完成時長)[^\d]*([\d\.]+s?)', page_text, re.IGNORECASE)
                    if comp_match:
                        comp_time = comp_match.group(1).replace(" ", "")

                    print(f"   ⏱️ 提取时间成功 -> 准备耗时: {prep_time}, 完成耗时: {comp_time}")
                except Exception as time_err:
                    print(f"   ⚠️ 提取时间信息时发生小错误，已跳过: {time_err}")
            # 2. 提取本地文档内容
                if is_text_only:
                    file_content = "【无原始文档，用户仅提供了纯文本提问，请仅根据问题本身评估回答是否准确且符合逻辑】"
                else:
                    file_content = read_file_content(file_path)

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
                        tester_exp,  # Tester Expectation
                        display_filename,  # 测试filename
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
                        timeout_status,
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
                #将等待时间缩短为 5 秒，找不到就赶紧刷新，不浪费时间
                home_btn = WebDriverWait(driver, 7).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "image-home"))
                )
                #使用 JS 强制点击，防止被网页的其他弹窗或动画层遮挡
                driver.execute_script("arguments[0].click();", home_btn)
                print("   ✅ 成功通过点击 Logo 返回")
            except Exception:
                # 当找不到 image-home 元素，或者点击失败时，直接请求网址
                print("   ⚠️ 未识别到主页 Logo，直接重新请求首页网址...")
                driver.get(home_url)

                # 留出一点时间让页面完全加载，准备迎接下一个文件
            time.sleep(1)
            close_popups(driver)
        # ==================================================

        except Exception as e:
            print(f"❌ [报错] 处理 {filename} 时发生错误: {e}")
            error_msg = str(e).lower()
            if "window already closed" in error_msg or "target window already closed" in error_msg or "no such window" in error_msg:
                print("🚨 侦测到浏览器窗口已被手动关闭，脚本放弃后续所有任务！")
                break

            # 如果中间发生报错，最稳妥的重置方式是直接重新请求首页网址
            driver.get(home_url)
            time.sleep(2)

    if not STOP_SCRIPT:
        print("📊 正在生成最终的 Summary 报告...")
        # 你的 input_excel_path 这里可能需要改成你实际生成的 evaluation_results.xlsx 的路径
        dynamic_csv_name = f"Summary_{base_testcase_name}_{timestamp}.csv"
        summary_dir = os.path.join(project_dir, f"Summaries_{env_name}")
        os.makedirs(summary_dir, exist_ok=True)
        output_csv = os.path.join(summary_dir, dynamic_csv_name)
        generate_summary_csv(excel_path, output_csv)
        print(f"✅ 汇总报告已生成: {output_csv}")
    if STOP_SCRIPT:
        print("\n🛑 任务已被手动中断！")
    else:
        print("\n✅ 文件夹内所有测试用例已全部运行完毕！")


if __name__ == "__main__":
    listener_thread = threading.Thread(target=listen_for_hotkey, daemon=True)
    listener_thread.start()
    # 配置两个网页的环境参数
    URL_1 = "https://customs-demo.poffices.ai/"
    ENV_NAME_1 = "customs"  # 将生成 Evaluation_Results_Site_A 文件夹

    URL_2 = "https://poffices.ai/"  # 替换为你的第二个网址
    ENV_NAME_2 = "poffices"  # 将生成 Evaluation_Results_Site_B 文件夹

    print("🚀 开始多线程并行执行自动化任务...")

    # 创建两个独立的线程，各自运行一个浏览器实例
    t1 = threading.Thread(target=run_automation, args=(URL_1, ENV_NAME_1))
    t2 = threading.Thread(target=run_automation, args=(URL_2, ENV_NAME_2))

    # 同时启动
    t1.start()
    t2.start()

    # 主线程在此等待，直到两个浏览器的任务都执行完毕
    t1.join()
    t2.join()

    print("👉 所有自动化任务均已结束。控制权已释放（浏览器将由 detach 属性保持开启）。")
    #print("等待手动按 ESC 键退出控制台...")

    # 维持控制台存活，直到按下 ESC
    while not STOP_SCRIPT:
        time.sleep(1)
