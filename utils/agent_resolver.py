import difflib
from dataclasses import dataclass
from typing import List, Callable, Optional

# 导入你原有的映射表
from tools.agent_mapping import AGENT_NAME_MAPPING_TC, TC_TO_ENG_MAPPING

# ================= 全局静态资源预加载 =================
# 动态生成所有合法的 Agent 名称全局列表（完全对齐你原来的逻辑，只在模块加载时运行一次，节省性能）
ALL_VALID_AGENTS = list(AGENT_NAME_MAPPING_TC.keys())
for tc_list in AGENT_NAME_MAPPING_TC.values():
    ALL_VALID_AGENTS.extend(tc_list)
ALL_VALID_AGENTS = list(set(ALL_VALID_AGENTS))

LOWER_AGENTS_MAPPING = {agent.lower(): agent for agent in ALL_VALID_AGENTS}


# ================= 数据返回模型 =================
@dataclass
class AgentResolutionResult:
    """封装 Agent 解析后的所有返回信息"""
    current_state: str  # "2" (通用模式) 或 "4" (指定 Agent 模式)
    final_target_agent: str  # 最终要在 Excel 中显示的 Agent 名称
    search_candidates: List[str]  # 传给网页自动化去点击的候选词列表
    is_tc_ui: bool  # 网页 UI 是否需要切到繁体中文


def resolve_agent_and_language(
        raw_agent: Optional[str],
        raw_language: Optional[str],
        log_func: Callable[[str], None] = print
) -> AgentResolutionResult:
    """
    核心业务逻辑：处理 Agent 的拼写纠错、大小写纠正、以及界面的多语言映射。

    :param raw_agent: 从 Excel 中读取的原始 Agent 名称
    :param raw_language: 从 Excel 中读取的 Selected Language
    :param log_func: 日志打印函数（兼容多线程的 safe_print 或单线程的 print）
    :return: AgentResolutionResult 对象
    """

    # 1. ==== 判断目标 UI 语言 ====
    tc_keywords = ["繁中", "繁体", "繁體", "traditional chinese", "tc", "zh-tw"]
    is_tc_ui = False
    if raw_language:
        is_tc_ui = any(keyword in str(raw_language).lower() for keyword in tc_keywords)

    # 2. ==== 核心状态与名称初始化 ====
    target_agent_str = str(raw_agent).strip() if raw_agent else ""
    current_state = "4" if target_agent_str else "2"

    # 3. ==== 拼写纠错与智能降级逻辑 (100% 对齐原版) ====
    if target_agent_str and target_agent_str not in ALL_VALID_AGENTS:
        target_agent_lower = target_agent_str.lower()

        # 第一步防禦：嘗試「無視大小寫的精確匹配」
        if target_agent_lower in LOWER_AGENTS_MAPPING:
            correct_name = LOWER_AGENTS_MAPPING[target_agent_lower]
            log_func(f"🔧 [大小寫修正]: 偵測到大小寫不標準 '{target_agent_str}'，已自動修正為 '{correct_name}'")
            target_agent_str = correct_name
        else:
            # 第二步防禦：模糊匹配 (cutoff=0.6)
            matches = difflib.get_close_matches(
                target_agent_lower, list(LOWER_AGENTS_MAPPING.keys()), n=1, cutoff=0.6
            )
            if matches:
                correct_name = LOWER_AGENTS_MAPPING[matches[0]]
                log_func(f"🔧 [智能糾錯]: 偵測到拼寫錯誤 '{target_agent_str}'，已自動糾正為合法的 '{correct_name}'")
                target_agent_str = correct_name
            else:
                log_func(
                    f"⚠️ [降級警告]: 拼写错误离谱，无法识别 Agent '{target_agent_str}'，自動降級為通用模式 (State 2)！")
                target_agent_str = ""
                current_state = "2"

    # 4. ==== 多语言 Agent 候选列表映射 (100% 对齐原版) ====
    search_candidates = []
    if target_agent_str:
        if is_tc_ui:
            # 【情况 A】UI 是繁中
            if target_agent_str in AGENT_NAME_MAPPING_TC:
                search_candidates = AGENT_NAME_MAPPING_TC[target_agent_str]
                log_func(f"🔄 [Agent 转换]: 繁中 UI 匹配到英文输入，载入候选列表 -> {search_candidates}")
            else:
                search_candidates = [target_agent_str]
                log_func(f"🎯 [Agent 保持]: 繁中 UI 匹配到繁中输入，直接搜索 -> '{target_agent_str}'")
        else:
            # 【情况 B】UI 是英文（默认）
            if target_agent_str in TC_TO_ENG_MAPPING:
                search_candidates = [TC_TO_ENG_MAPPING[target_agent_str]]
                log_func(f"🔄 [Agent 转换]: 英文 UI 匹配到繁中输入，已自动转为英文 -> '{search_candidates[0]}'")
            else:
                search_candidates = [target_agent_str]
                log_func(f"🎯 [Agent 保持]: 英文 UI 匹配到英文输入，直接搜索 -> '{target_agent_str}'")

    return AgentResolutionResult(
        current_state=current_state,
        final_target_agent=target_agent_str,
        search_candidates=search_candidates,
        is_tc_ui=is_tc_ui
    )
