import time
import re
import threading
from typing import Callable, Optional, List, Tuple
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webdriver import WebDriver

# 导入原有的工具函数
from tools.state_manager import execute_state, handle_post_send
from utils.common import get_react_time_extraction_js, close_popups
from utils.signals import STOP_EVENT

class WebInteractor:
    """
    网页交互核心类。
    负责所有与 Selenium 相关的 DOM 操作、JS 注入、等待监控等。
    """

    def __init__(self, driver: WebDriver, log_func: Callable[[str], None] = print):
        self.driver = driver
        self.log_func = log_func

    def check_browser_alive(self):
        """检查浏览器是否被手动关闭"""
        try:
            _ = self.driver.window_handles
        except Exception as e:
            raise Exception("window already closed") from e

    def switch_language(self, is_tc_ui: bool):
        """切换 UI 语言 (100% 对齐原版)"""
        if is_tc_ui:
            self.log_func("🌐 [语言切换]: 正在切换为 繁中...")
            lang_container = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "language-container")))
            self.driver.execute_script("arguments[0].click();", lang_container)
            time.sleep(1)
            tc_option = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located(
                (By.XPATH, "//a[contains(@class, 'language-option') and text()='繁中']")))
            self.driver.execute_script("arguments[0].click();", tc_option)
            time.sleep(1)
        else:
            self.log_func("🌐 [语言保持]: 目标语言为英文或空白，无需切换，直接保持默认英文状态。")
            """
            self.log_func("🌐 [语言切换]: 正在切换为 英文 (默认)...")
            lang_container = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "language-container")))
            self.driver.execute_script("arguments[0].click();", lang_container)
            time.sleep(1)
            eng_option = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located(
                (By.XPATH, "//a[contains(@class, 'language-option') and text()='ENG']")))
            self.driver.execute_script("arguments[0].click();", eng_option)
            time.sleep(1)
            """

    def reset_language_to_english(self):
        """网页返回前重置为英文 (防状态残留)"""
        try:
            self.log_func("🌐 [语言重置]: 正在准备返回前重置为英文...")
            globe_icon = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "language-icon")))
            self.driver.execute_script("arguments[0].click();", globe_icon)
            time.sleep(1)
            eng_reset_option = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(@class, 'language-option') and text()='ENG']")))
            self.driver.execute_script("arguments[0].click();", eng_reset_option)
            self.log_func("✅ 成功点击 ENG 重置语言")
        except Exception as reset_lang_err:
            self.log_func(f"⚠️ 重置语言步骤失败 (可能页面已改变)，将继续执行: {reset_lang_err}")

    def upload_file_if_exists(self, file_path: Optional[str]):
        """上传文件 (如果存在)"""
        if file_path:
            file_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
            file_input.send_keys(file_path)
            time.sleep(1)
        else:
            self.log_func("📄 [跳过]: 纯文本模式，无需上传文件。")

    def inject_question(self, question_text: str):
        """JS 强行注入问题文本 (完美兼容 React/Vue 底层值追踪)"""
        safe_text = question_text if question_text else ""

        try:
            text_area = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "background-info")))
            self.driver.execute_script("arguments[0].value = '';", text_area)
            self.driver.execute_script("arguments[0].value = arguments[1];", text_area, safe_text)
            self.driver.execute_script("""
                var input = arguments[0];
                var lastValue = input.value;
                input.value = arguments[1];
                var event = new Event('input', { bubbles: true });
                var tracker = input._valueTracker;
                if (tracker) { tracker.setValue(lastValue); }
                input.dispatchEvent(event);
            """, text_area, question_text)
            self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", text_area)
            self.log_func("✅ [JS 注入成功]: 已成功在后台写入问题文本。")
            time.sleep(1)
        except Exception as input_err:
            self.log_func(f"❌ 输入问题文本时发生异常: {input_err}")
            raise input_err

    def apply_agent_and_submit(self, current_state: str, search_candidates: List[str], target_agent_raw: str) -> str:
        """点击齿轮、应用 Agent 并提交 (包含对齐的保底降级机制)"""
        self.log_func("🔍 [通用步骤]: 正在点击设置齿轮图标...")
        gear_btn = WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'bi-gear-fill')]")))
        self.driver.execute_script("arguments[0].click();", gear_btn)
        WebDriverWait(self.driver, 5).until(
            EC.visibility_of_element_located((By.ID, "global-settings-apply-btn")))

        final_agent_used = target_agent_raw
        if search_candidates:
            agent_success = False
            for candidate in search_candidates:
                try:
                    self.log_func(f"👉 正在尝试选中 Agent: '{candidate}' ...")
                    execute_state(self.driver, current_state, candidate)
                    agent_success = True
                    self.log_func(f"✅ 成功找到并应用 Agent: '{candidate}'")
                    break
                except Exception:
                    self.log_func(f"⚠️ 当前网站未找到 '{candidate}'，准备尝试下一个候选词...")

            if not agent_success:
                self.log_func(f"⚠️ [保底机制触发]: 候选 Agent {search_candidates} 均无法找到！已自动降级为 State 2。")
                current_state = "2"
                final_agent_used = f"{target_agent_raw} (未找到，降级为 State 2)"
                execute_state(self.driver, current_state)
        else:
            self.log_func("ℹ️ 目标 Agent 为空，直接应用通用模式 (State 2)...")
            execute_state(self.driver, current_state)

        self.log_func("🔍 [通用步骤]: 面板配置完毕，正在点击 Apply Settings...")
        apply_btn = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "global-settings-apply-btn")))
        self.driver.execute_script("arguments[0].click();", apply_btn)
        WebDriverWait(self.driver, 5).until(
            EC.invisibility_of_element_located((By.ID, "global-settings-apply-btn")))

        self.log_func("🚀 [通用步骤]: 正在点击发送/生成按钮...")
        submit_btn = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'btn-enter')]")))
        self.driver.execute_script("arguments[0].click();", submit_btn)

        handle_post_send(self.driver, current_state)
        self.log_func("✅ 已成功进入自动生成等待环节...")

        return final_agent_used

    def monitor_generation(self, stop_event: threading.Event) -> str:
        """智能监控 AI 生成进度 (300秒防死循环)"""
        self.log_func("⏳ 正在智能监控 AI 生成进度...")
        timeout_status = "No"
        try:
            stop_event.wait(timeout=5)  # 强制等待 5 秒前端动画
            last_length = 0
            stable_count = 0

            for _ in range(300):
                if stop_event.is_set():
                    self.log_func("🛑 收到中止指令，立即停止网页监控！")
                    break

                self.check_browser_alive()
                time.sleep(1)

                current_length = 0
                try:
                    previews = self.driver.find_elements(By.ID, "preview")
                    for p in reversed(previews):
                        if p.is_displayed():
                            actual_text = p.get_attribute("innerText") or p.get_attribute("textContent") or ""
                            current_length = len(actual_text.strip())
                            break
                except Exception:
                    pass

                if current_length > 0 and current_length == last_length:
                    stable_count += 1
                else:
                    stable_count = 0
                last_length = current_length

                if stable_count >= 5:  # 连续 5 秒稳定
                    self.log_func(f"✅ 回答文本已连续 5 秒无变化，判定生成彻底完成！最终字数: {current_length}")
                    break
            else:
                if not stop_event.is_set():
                    self.log_func("⚠️ 警告：监控达到 300 秒上限，生成总时间超时！")
                    timeout_status = "yes (总时间超时)"
        except Exception as wait_error:
            # 【绝不吞噬关窗异常，直接往外抛，让 pipeline 捕获】
            error_msg = str(wait_error).lower()
            if "window already closed" in error_msg or "no such window" in error_msg or "target window already closed" in error_msg:
                raise wait_error
            self.log_func(f"⚠️ 智能监控发生异常，强制继续执行: {wait_error}")
            time.sleep(2)
        return timeout_status

    def extract_answer(self, stop_event: threading.Event) -> Tuple[str, str]:
        """提取最新回答与链接 (60秒等待)"""
        self.log_func("📥 正在提取生成的回答...")
        answer_text = ""
        current_page_url = self.driver.current_url
        self.log_func(f"🔗 已抓取当前回答 shared link: {current_page_url}")

        for _ in range(60):
            if stop_event.is_set():
                break
            try:
                self.check_browser_alive()
                previews = self.driver.find_elements(By.ID, "preview")
                for p in reversed(previews):
                    actual_content = p.get_attribute("innerText") or p.get_attribute("textContent") or ""
                    if len(actual_content.strip()) > 0:
                        answer_text = actual_content
                        break
                if answer_text:
                    break
            except Exception as e:
                # 【拦截浏览器关闭信号并立刻跳出】
                error_msg = str(e).lower()
                if "window already closed" in error_msg or "no such window" in error_msg or "target window already closed" in error_msg:
                    self.log_func("🚨 提取内容时侦测到浏览器已关闭，终止当前任务！")
                    return "BROWSER_CLOSED", current_page_url
                pass  # 其他查找元素偶发的报错依然忽略
            time.sleep(1)
        if stop_event.is_set():
            self.log_func("🛑 收到中止指令，立即中断提取动作！")
            return "", current_page_url
        if not answer_text:
            self.log_func("⚠️ 警告：等待 60 秒后依然未抓取到内容！")
            answer_text = "提取文本失败/为空"
        else:
            self.log_func(f"✅ 成功提取到回答，长度: {len(answer_text)} 字符")

        return answer_text, current_page_url

    def extract_react_times(self, stop_event: threading.Event) -> Tuple[str, str]:
        """注入核弹级 JS 提取 React 渲染时间"""
        prep_time, comp_time = "N/A", "N/A"
        try:
            if stop_event.wait(timeout=3):
                return "N/A", "N/A"
            times_info = self.driver.execute_script(get_react_time_extraction_js())
            if times_info:
                prep_time = times_info.get("prep", "N/A")
                comp_time = times_info.get("comp", "N/A")
                debug_info = times_info.get("debug", "")

                if prep_time == "N/A":
                    html = self.driver.execute_script("return document.body.innerHTML;")
                    html_matches = re.findall(r'<number-flow[^>]*?(?:value|aria-valuenow)=["\']?([0-9.]+)["\']?', html,
                                              re.IGNORECASE)
                    if len(html_matches) >= 1:
                        prep_time = html_matches[0] + "s"
                        debug_info += " | 触发底层HTML正则兜底"
                    if len(html_matches) >= 2:
                        comp_time = html_matches[1] + "s"

                self.log_func(f"⏱️ 提取时间 -> 准备耗时: {prep_time}, 完成耗时: {comp_time} (底层诊断: {debug_info})")
            else:
                self.log_func("⏱️ 提取失败：JS 脚本未返回任何数据。")
        except Exception as e:
            self.log_func(f"⚠️ 提取时间发生代码异常: {e}")

        return prep_time, comp_time

    def return_to_home(self, home_url: str):
        """安全返回首页重置状态"""
        try:
            self.log_func("🏠 正在尝试通过点击 Logo 返回首页...")
            home_btn = WebDriverWait(self.driver, 7).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "image-home")))
            self.driver.execute_script("arguments[0].click();", home_btn)
            self.log_func("✅ 成功通过点击 Logo 返回")
        except Exception:
            self.log_func("⚠️ 未识别到主页 Logo，直接重新请求首页网址...")
            try:
                self.driver.execute_script("window.sessionStorage.clear();")
            except:
                pass
            self.driver.get(home_url)
        time.sleep(1)
        close_popups(self.driver, log_func=self.log_func)
