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
from concurrent.futures import ThreadPoolExecutor
from threading import Lock
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


def process_single_task(i, question_text, target_agent, filename, target_language, test_dir, excel_path,
                        excel_write_lock, USERNAME, PASSWORD, HOME_URL):
    global STOP_SCRIPT

    if STOP_SCRIPT:
        print(f"🛑 [任务 {i + 1}] 收到中止指令，停止处理。")
        return

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
            else:
                print(f"\n❌ [跳过任务] 找不到文件 '{filename}' 且没有 Request，无法继续，跳过此条任务！")
                return
    else:
        # 如果 Excel 里没写 filename，再双重确认下有没有 request
        if not question_text or not str(question_text).strip():
            print(f"\n❌ [跳过任务] Request 和 Filename 同时为空，跳过此条任务！")
            return

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
        if target_agent_raw in AGENT_NAME_MAPPING_TC:
            search_candidates = AGENT_NAME_MAPPING_TC[target_agent_raw]
            print(f"   🔄 [Agent 转换]: 繁中 UI 匹配到英文输入，载入候选列表 -> {search_candidates}")
        else:
            search_candidates = [target_agent_raw]
            print(f"   🎯 [Agent 保持]: 繁中 UI 匹配到繁中输入，直接搜索 -> '{target_agent_raw}'")
    else:
        if target_agent_raw in TC_TO_ENG_MAPPING:
            search_candidates = [TC_TO_ENG_MAPPING[target_agent_raw]]
            print(f"   🔄 [Agent 转换]: 英文 UI 匹配到繁中输入，已自动转为英文 -> '{search_candidates[0]}'")
        else:
            search_candidates = [target_agent_raw]
            print(f"   🎯 [Agent 保持]: 英文 UI 匹配到英文输入，直接搜索 -> '{target_agent_raw}'")

    # ==== 打印信息 ====
    print(f"\n--- 进度: {i + 1} ---")
    print(f"❓ 输入问题: {question_text if question_text else '【无文本，仅上传文件】'}")
    print(f"🤖 目标Agent: {target_agent if target_agent and str(target_agent).strip() else '【未指定，自动使用状态 2】'}")
    print(f"⚙️ 最终执行状态: {CURRENT_STATE}")

    if not is_text_only:
        print(f"📄 准备上传文件: {actual_file_name}")
    else:
        print(f"📄 纯文本/降级模式，无需上传文件。")

    driver = None
    try:
        # ================= 启动专属浏览器 =================
        print(f"[{i + 1}] 🔍 尝试启动首选浏览器：Google Chrome...")
        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("detach", True)
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        except Exception as e_chrome:
            print(f"[{i + 1}] ⚠️ Chrome 启动失败: {e_chrome}")
            print(f"[{i + 1}] 🔄 正在尝试启动备用浏览器：Microsoft Edge...")
            edge_options = EdgeOptions()
            edge_options.add_experimental_option("detach", True)
            driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)

        driver.get(HOME_URL)
        time.sleep(1)

        # ================= 执行自动登录 =================
        print(f"[{i + 1}] 🔑 正在执行自动登录...")
        driver.find_element(By.ID, "nav-login").click()
        time.sleep(1)
        driver.find_element(By.ID, "sso-username").send_keys(USERNAME)
        driver.find_element(By.ID, "sso-password").send_keys(PASSWORD)
        driver.find_element(By.CLASS_NAME, "sso-submit-btn").click()
        time.sleep(1)
        close_popups(driver)
        print(f"[{i + 1}] ✅ 登录动作已提交，等待页面刷新...")
        time.sleep(1)
        close_popups(driver)

        # --- 网页自动化交互逻辑 ---
        if target_language and ("繁中" in str(target_language) or "繁体" in str(target_language)):
            print("   🌐 [语言切换]: 正在切换为 繁中...")
            lang_container = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "language-container")))
            driver.execute_script("arguments[0].click();", lang_container)
            time.sleep(1)
            tc_option = WebDriverWait(driver, 3).until(EC.presence_of_element_located(
                (By.XPATH, "//a[contains(@class, 'language-option') and text()='繁中']")))
            driver.execute_script("arguments[0].click();", tc_option)
            time.sleep(1.5)
        else:
            print("   🌐 [语言切换]: 正在切换为 英文 (默认)...")
            lang_container = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "language-container")))
            driver.execute_script("arguments[0].click();", lang_container)
            time.sleep(1)
            eng_option = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'language-option') and text()='ENG']")))
            driver.execute_script("arguments[0].click();", eng_option)
            time.sleep(1)

        if not is_text_only and file_path:
            file_input = driver.find_element(By.XPATH, "//input[@type='file']")
            file_input.send_keys(file_path)
            time.sleep(1)
        else:
            print("   📄 [跳过]: 纯文本模式，无需上传文件。")

        text_area = driver.find_element(By.ID, "background-info")
        text_area.clear()
        text_area.send_keys(question_text)
        time.sleep(1)

        print("   🔍 [通用步骤]: 正在点击设置齿轮图标...")
        gear_btn = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'bi-gear-fill')]")))
        driver.execute_script("arguments[0].click();", gear_btn)
        time.sleep(0.5)

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
            raise Exception(f"所有候选 Agent {search_candidates} 均无法找到！报错信息: {last_err}")

        print("   🔍 [通用步骤]: 面板配置完毕，正在点击 Apply Settings...")
        apply_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "global-settings-apply-btn")))
        driver.execute_script("arguments[0].click();", apply_btn)
        time.sleep(1)
        print("   🚀 [通用步骤]: 正在点击发送/生成按钮...")
        submit_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'btn-enter')]")))
        driver.execute_script("arguments[0].click();", submit_btn)

        handle_post_send(driver, CURRENT_STATE)
        print("   ✅ 已成功进入自动生成等待环节...")

        print("⏳ 正在智能监控 AI 生成进度...")
        timeout_status = "No"
        time.sleep(5)
        last_length = 0
        stable_count = 0
        max_wait_loops = 300
        required_stable_seconds = 5

        for _ in range(max_wait_loops):
            if STOP_SCRIPT:
                print("🛑 收到中止指令，立即停止网页监控！")
                break
            try:
                _ = driver.window_handles
            except Exception:
                print("\n🚨 监控期间侦测到浏览器窗口被关闭，立即退出！")
                break
            time.sleep(1)
            current_length = 0
            try:
                previews = driver.find_elements(By.ID, "preview")
                for p in reversed(previews):
                    if p.is_displayed():
                        current_length = len(p.text.strip())
                        break
            except Exception:
                pass

            if current_length > 0 and current_length == last_length:
                stable_count += 1
            else:
                stable_count = 0
            last_length = current_length

            if stable_count >= required_stable_seconds:
                print(
                    f"✅ 回答文本已连续 {required_stable_seconds} 秒无变化，判定生成彻底完成！最终字数: {current_length}")
                break
        else:  # <--- 新增：注意与上面的 for 对齐
            if not STOP_SCRIPT:
                print(f"[{i + 1}] ⚠️ 警告：监控达到 300 秒上限，生成总时间超时！")
                timeout_status = "yes (总时间超时)"
        if STOP_SCRIPT:
            return

        current_page_url = driver.current_url
        print(f"🔗 已抓取当前回答shared link: {current_page_url}")
        print("📥 正在提取生成的回答并读取原文件...")

        def get_valid_preview(d):
            try:
                _ = d.window_handles
            except Exception:
                return "BROWSER_CLOSED"
            previews = d.find_elements(By.ID, "preview")
            for p in reversed(previews):
                if p.is_displayed() and len(p.text.strip()) > 0:
                    return p.text
            return False

        try:
            answer_text = WebDriverWait(driver, 60).until(get_valid_preview)
            if answer_text == "BROWSER_CLOSED":
                print("🚨 提取内容时侦测到浏览器已关闭，终止当前任务！")
                return
        except Exception:
            answer_text = ""

        if not answer_text or not answer_text.strip():
            print("   ⚠️ 警告：等待 60 秒后依然未抓取到内容！")
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
            time.sleep(1)
            page_text = driver.find_element(By.TAG_NAME, "body").text
            prep_match = re.search(r'(?:Time of preparation|準備時長:)[^\d]*([\d\.]+s?)', page_text, re.IGNORECASE)
            if prep_match: prep_time = prep_match.group(1).replace(" ", "")
            comp_match = re.search(r'(?:Time of completion|完成時長)[^\d]*([\d\.]+s?)', page_text, re.IGNORECASE)
            if comp_match: comp_time = comp_match.group(1).replace(" ", "")
            print(f"   ⏱️ 提取时间成功 -> 准备耗时: {prep_time}, 完成耗时: {comp_time}")
        except Exception as time_err:
            print(f"   ⚠️ 提取时间信息时发生小错误，已跳过: {time_err}")

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
        evaluation_text = eval_results.get("evaluation", "Unknown")
        ref_link = eval_results.get("reference_link", "N/A")
        doc_contain = eval_results.get("document_contain_citations", "None")

        # ================= 加上线程锁写入 Excel =================
        with excel_write_lock:
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                short_answer = answer_text[:100] + "..." if len(answer_text) > 100 else answer_text
                display_filename = filename if filename else ""
                display_target_language = target_language if target_language else "N/A"

                # 准备好要写入的一整行数据
                row_data = [
                    i + 1, question_text, tester_exp, display_filename, display_target_language,
                    input_lang, output_lang, lang_status, short_answer, current_page_url,
                    evaluation_text, target_agent if target_agent else "未指定", ref_link, doc_contain,
                    prep_time, comp_time
                ]

                # 精准占位写入：因为 Excel 第1行是表头，所以第 i 个任务应该写在第 i + 2 行
                target_row = i + 2
                for col_index, value in enumerate(row_data, start=1):
                    ws.cell(row=target_row, column=col_index, value=value)
                wb.save(excel_path)
                print(f"[{i + 1}] ✅ 评价结果已成功写入 Excel 文件。")
            except Exception as excel_err:
                print(f"[{i + 1}] ❌ 写入 Excel 时失败: {excel_err}")

    except Exception as e:
        print(f"[{i + 1}] ❌ 处理 {filename} 时发生错误: {e}")
        error_msg = str(e).lower()
        if "window already closed" in error_msg or "target window already closed" in error_msg or "no such window" in error_msg:
            print(f"[{i + 1}] 🚨 侦测到浏览器窗口已被手动关闭，终止当前子任务！")
    finally:
        # 核心：多线程环境下，单个任务跑完必须关闭独立浏览器释放内存
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


def run_automation():
    global STOP_SCRIPT
    # 开启后台键盘监听线程
    listener_thread = threading.Thread(target=listen_for_hotkey, daemon=True)
    listener_thread.start()

    # ================= 动态路径配置（替代写死的绝对路径） =================
    USERNAME = "HenryHONG"
    PASSWORD = "12345678"
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
    valid_extensions = ('.pdf', '.docx', '.csv', '.txt')
    test_files = [f for f in os.listdir(test_dir) if f.lower().endswith(valid_extensions)]
    test_files.sort()

    if len(test_files) != len(questions):
        print(f"[警告] 文件夹中的文件数量 ({len(test_files)}) 与 Excel 中的问题数量 ({len(questions)}) 不匹配！")

    # ================= 初始化 Excel 结果文件 =================
    base_testcase_name = os.path.splitext(os.path.basename(input_excel_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dynamic_filename = f"evaluation_results_{base_testcase_name}_{timestamp}.xlsx"

    # 定义并创建 Evaluation_Results 文件夹
    excel_dir = os.path.join(project_dir, "Evaluation_Results")
    os.makedirs(excel_dir, exist_ok=True)

    # 修改：将 excel_path 指向新的子文件夹
    excel_path = os.path.join(excel_dir, dynamic_filename)

    if not os.path.exists(excel_path):
        wb = openpyxl.Workbook()
        wb.properties.creator = "Henry HONG (洪伟恒)"
        wb.properties.description = "Authored by HONGWEIHENG. Tel: 17722596827"
        ws = wb.active
        ws.title = "Evaluation Results"
        ws.append(["label", "Request", "Tester Expectation", "filename", "Selected Language", "Input Language",
                   "Output Language", "Language Overall Status", "answer", "shared link", "DeepSeek评价内容",
                   "Selected agent", "Reference Link", "Document Contain[1][2][3]", "Preparation Time",
                   "Completion Time","Timeout_States"])
        wb.save(excel_path)
        print(f"📊 已创建评价结果 Excel 文件: {excel_path}")
    # ===============================================================

    # ==================== 并发执行核心区 ====================
    excel_write_lock = Lock()
    HOME_URL = "https://customs-demo.poffices.ai/"

    print("\n🚀 启动并发模式处理任务...")
    # max_workers=3 表示同时开 3 个网页跑，你可以根据电脑性能调整
    with ThreadPoolExecutor(max_workers=5) as executor:
        for i in range(len(questions)):
            if STOP_SCRIPT:
                break
            executor.submit(
                process_single_task,
                i, questions[i], selected_agents[i], target_filenames[i], selected_languages[i],
                test_dir, excel_path, excel_write_lock, USERNAME, PASSWORD, HOME_URL
            )
            time.sleep(2)  # 错峰启动浏览器，防止 CPU 瞬间拉满
    # ===============================================================

    # 等所有线程都跑完后，才会执行下面的报告生成逻辑
    if not STOP_SCRIPT:
        print("📊 正在生成最终的 Summary 报告...")
        dynamic_csv_name = f"Summary_{base_testcase_name}_{timestamp}.csv"
        summary_dir = os.path.join(project_dir, "Summaries")
        os.makedirs(summary_dir, exist_ok=True)
        output_csv = os.path.join(summary_dir, dynamic_csv_name)
        print(f"✅ 汇总报告已生成: {output_csv}")

    if STOP_SCRIPT:
        print("\n🛑 任务已被手动中断！")
    else:
        print("\n✅ 文件夹内所有测试用例已全部运行完毕！")

    print("👉 自动化流程结束，控制台即将退出。")
    #input("按 ESC 键退出当前控制台窗口...")
if __name__ == "__main__":
    run_automation()
