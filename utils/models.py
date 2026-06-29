from dataclasses import dataclass
from typing import Optional

@dataclass
class TaskInput:
    """
    统一封装单条任务的输入数据。
    完全对齐你原来的: questions[i], selected_agents[i], target_filenames[i], selected_languages[i]
    """
    index: int                  # 对应 Excel 的行号/索引 (i)
    question_text: str          # 对应 Request
    target_agent: str           # 对应 Selected agent
    filename: str               # 对应 filename
    target_language: str        # 对应 Selected Language

@dataclass
class TaskResult:
    """
    统一封装单条任务的执行结果，准备用于写入 Excel。
    完全对齐你原脚本中准备写入 ws.append([...]) 的那一长串变量。
    """
    crash_reason: str = "No"                 # 默认无崩溃
    tester_expectation: str = "Unknown"
    input_language: str = "N/A"
    output_language: str = "N/A"
    language_status: str = "Unknown"
    answer_text: str = ""                    # 生成的回答或报错信息
    shared_link: str = "N/A"                 # 当前页面 URL
    evaluation_text: str = "Unknown"         # DeepSeek 评价或报错复用
    actual_agent_used: str = "未指定"        # 最终实际使用的 Agent
    reference_link: str = "N/A"
    document_contain: str = "None"
    prep_time: str = "N/A"
    comp_time: str = "N/A"
    timeout_status: str = "No"               # 超时状态
