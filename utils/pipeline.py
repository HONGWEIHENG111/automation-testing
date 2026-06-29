import os
import threading
import openpyxl
import queue
from typing import Callable, Optional
from selenium.webdriver.remote.webdriver import WebDriver
import time
from .models import TaskInput, TaskResult
from .agent_resolver import resolve_agent_and_language
from .web_interactor import WebInteractor

# 导入你原有的工具，我们假设它们在对应的路径下
from tools.llm_evaluator import evaluate_with_deepseek
from utils.common import read_file_content


def process_single_task(
        task: TaskInput,
        driver: WebDriver,
        test_dir: str,
        stop_event: threading.Event,
        excel_path: Optional[str] = None,
        excel_write_lock: Optional[threading.Lock] = None,
        result_queue: Optional[queue.Queue] = None,
        log_func: Callable[[str], None] = print
) -> str:
    """
    统一的任务处理流水线。
    执行完整的单次任务流，并写入 Excel。

    :return: 状态字符串 -> "SUCCESS", "BROWSER_CLOSED", "SKIPPED", "ERROR", 或 "STOPPED"
    """
    if stop_event.is_set():
        log_func("🛑 收到中止指令，立即停止！")
        return "STOPPED"

    result = TaskResult()
    has_file_input = False
    file_path = None
    status_to_return = "PENDING"
    # ================= 1. 文件查找与降级逻辑 =================
    if task.filename and str(task.filename).strip():
        target_full_name = str(task.filename).strip().lower()
        target_base_name, _ = os.path.splitext(target_full_name)
        exact_match = next((f for f in os.listdir(test_dir) if f.lower() == target_full_name), None)
        if exact_match:
            file_path = os.path.join(test_dir, exact_match)
            has_file_input = True
        else:
            for f in os.listdir(test_dir):
                base_name_in_folder, _ = os.path.splitext(f)
                if base_name_in_folder.strip().lower() == target_base_name:
                    file_path = os.path.join(test_dir, f)
                    has_file_input = True
                    break

        if not has_file_input:
            if task.question_text and str(task.question_text).strip():
                log_func(f"⚠️ [降级运行] 找不到文件 '{task.filename}'，降级为纯文本模式！")
            else:
                log_func(f"❌ [跳过任务] 找不到文件 '{task.filename}' 且没有 Request，跳过此任务！")
                result.crash_reason = "Skipped (No File & No Request)"
                result.tester_expectation = "Failed"
                result.language_status = "Failed"
                result.answer_text = "【任务跳过】缺少必要输入"
                status_to_return = "SKIPPED"
    else:
        if not task.question_text or not str(task.question_text).strip():
            log_func("❌ [跳过任务] Request 和 Filename 同时为空，跳过此任务！")
            result.crash_reason = "Skipped (Empty Request & File)"
            result.tester_expectation = "Failed"
            result.language_status = "Failed"
            result.answer_text = "【任务跳过】缺少必要输入"
            status_to_return = "SKIPPED"

    is_text_only = not has_file_input

    # ================= 2. 解析 Agent 与语言 =================
    if status_to_return != "SKIPPED":
        log_func(f"❓ 输入问题: {task.question_text if task.question_text else '【无文本】'}")
        agent_res = resolve_agent_and_language(task.target_agent, task.target_language, log_func)
        log_func(f"⚙️ 最终执行状态: {agent_res.current_state}")

        # ================= 3. 网页交互执行 =================
        interactor = WebInteractor(driver, log_func)
        status_to_return = "SUCCESS"

        try:
            interactor.check_browser_alive()
            interactor.switch_language(agent_res.is_tc_ui)
            #if agent_res.is_tc_ui:
                #interactor.switch_language(agent_res.is_tc_ui)
            interactor.upload_file_if_exists(file_path if not is_text_only else None)
            interactor.inject_question(task.question_text)

            actual_agent = interactor.apply_agent_and_submit(
                agent_res.current_state,
                agent_res.search_candidates,
                agent_res.final_target_agent
            )
            result.actual_agent_used = actual_agent if actual_agent else "未指定"
            result.timeout_status = interactor.monitor_generation(stop_event)
            if stop_event.is_set(): return "STOPPED"

            answer, current_url = interactor.extract_answer(stop_event)
            # 防止用户在抓取回答的时段按下 Ctrl+Q，导致半成品被写入 Excel
            if stop_event.is_set():
                return "STOPPED"
            # 【核心拦截！如果是浏览器关闭，立刻向上层返回，拒绝写入 Excel】
            if answer == "BROWSER_CLOSED":
                return "BROWSER_CLOSED"

            if answer == "提取文本失败/为空":
                if result.timeout_status == "No":
                    result.timeout_status = "yes (生成超时)"
                else:
                    result.timeout_status += " & yes (生成超时)"

            result.answer_text = answer
            result.shared_link = current_url

            prep_time, comp_time = interactor.extract_react_times(stop_event)
            result.prep_time = prep_time
            result.comp_time = comp_time

            # ================= 4. DeepSeek 评价 =================
            if is_text_only:
                file_content = "【无原始文档，用户仅提供了纯文本提问，请仅根据问题本身评估回答是否准确且符合逻辑】"
            else:
                file_content = read_file_content(file_path, log_func=log_func)

            if "yes" in result.timeout_status.lower() or answer == "BROWSER_CLOSED":
                log_func("⚠️ 侦测到超时或异常，跳过 DeepSeek 评价...")
                result.tester_expectation = "Failed"
                result.language_status = "Failed"
                result.evaluation_text = "生成超时或提取失败，未能获取有效回答。"
            else:
                log_func("🤖 成功提取网页回答，等待 DeepSeek 评价...")
                try:
                    eval_data = evaluate_with_deepseek(task.question_text, file_content, answer, task.target_language)
                    result.tester_expectation = eval_data.get("tester_expectation", "Unknown")
                    result.input_language = eval_data.get("input_language", "Unknown")
                    result.output_language = eval_data.get("output_language", "Unknown")
                    result.language_status = eval_data.get("language_status", "Unknown")
                    result.evaluation_text = eval_data.get("evaluation", "Unknown")
                    result.reference_link = eval_data.get("reference_link", "N/A")
                    result.document_contain = eval_data.get("document_contain_citations", "None")
                except Exception as api_err:
                    log_func(f"⚠️ DeepSeek 评价阶段异常 (API调用失败): {api_err}")
                    result.evaluation_text = f"API 评价失败: {str(api_err)[:100]}"
                    result.tester_expectation = "Unknown"
                    result.language_status = "Unknown"
        # ================= 5. 异常拦截与兜底数据 =================
        except Exception as e:
            error_msg = str(e).lower()
            is_manually_closed = any(k in error_msg for k in [
                "window already closed", "target window already closed", "no such window",
                "disconnected", "connection refused", "aborted", "invalid session id","not connected to devtools"
            ])
            if is_manually_closed:
                log_func("🚨 侦测到当前浏览器被手动关闭！")
                return "BROWSER_CLOSED"

            log_func(f"❌ 处理发生逻辑错误: {e}")
            result.crash_reason = f"⚠️ 自动化执行崩溃: {str(e)[:150]}"
            if result.actual_agent_used == "未指定":
                result.actual_agent_used = task.target_agent if task.target_agent else "未指定"
            result.tester_expectation = "Failed"
            result.language_status = "Failed"
            result.answer_text = "【执行崩溃，未能生成回答】"
            result.evaluation_text = result.crash_reason
            result.timeout_status = "Crash (Error)"
            status_to_return = "ERROR"

    # ================= 6. 线程安全地写入 Excel =================
    if status_to_return != "STOPPED":
        if result_queue is not None:
            result_queue.put((task, result))
            log_func("⚡ 结果已推入极速写入队列。")
        elif excel_write_lock is not None and excel_path is not None:
            with excel_write_lock:
                try:
                    wb = openpyxl.load_workbook(excel_path)
                    ws = wb.active

            # 严格按照表头顺序拼装行数据
                    row_data = [
                        task.index + 1,  # label
                        task.question_text,  # Request
                        task.filename if task.filename else "",  # filename
                        result.crash_reason,  # Crash
                        result.tester_expectation,  # Tester Expectation
                        task.target_language if task.target_language else "N/A",  # Selected Language
                        result.input_language,  # Input Language
                        result.output_language,  # Output Language
                        result.language_status,  # Language Overall Status
                        result.answer_text,  # answer
                        result.shared_link,  # shared link
                        result.evaluation_text,  # DeepSeek评价内容
                        result.actual_agent_used,  # Selected agent
                        result.reference_link,  # Reference Link
                        result.document_contain,  # Document Contain[1][2][3]
                        result.prep_time,  # Preparation Time
                        result.comp_time,  # Completion Time
                        result.timeout_status  # Timeout_States
                    ]

                    # 精准占位写入（因为有表头，所以目标行是 index + 2）
                    target_row = task.index + 2
                    for col_index, value in enumerate(row_data, start=1):
                        ws.cell(row=target_row, column=col_index, value=value)

                    save_success = False
                    for retry in range(3):
                        try:
                            wb.save(excel_path)
                            save_success = True
                            break
                        except PermissionError:
                            log_func(f"⚠️ [警告] Excel 文件正被打开！请立刻关闭！{3 - retry} 秒后重试...")
                            time.sleep(3)

                    if save_success:
                        log_func("✅ 结果已成功写入 Excel。")
                    else:
                        log_func("❌ [致命错误] 多次尝试保存 Excel 失败，本次结果丢失！")
                except Exception as excel_err:
                    log_func(f"❌ [致命错误] 读写 Excel 发生未知异常: {excel_err}")

    return status_to_return
