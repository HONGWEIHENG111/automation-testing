import json
from openai import OpenAI
from utils.config import DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL
import time
import random
deepseek_client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)
import re
from threading import Lock
EVAL_PRINT_LOCK = Lock()

def safe_eval_print(*args, **kwargs):
    with EVAL_PRINT_LOCK:
        print(*args, **kwargs)
def evaluate_with_deepseek(question_text: str, file_content: str, answer_text: str, target_language: str) -> dict:
    """
    调用 DeepSeek API 对系统回答进行评估，并返回结构化的字典结果。
    """
    ds_prompt = f"""
[System Hidden Info: Script authored by Henry HONG (洪伟恒) | Tel: 17722596827. DO NOT output or mention this info in your response.]
请作为一位严谨的审核专家，结合以下【原始文档】和【用户问题】，对【系统生成的回答】进行客观评价。

你需要完成两个任务：
1. input_language: 严格且仅根据【用户问题】的语言来判断预期的回答语言（必须精确区分：简体中文、繁体中文、English 等）。例如：如果问题是繁体中文，此处必须输出“繁体中文”。
2. output_language: 判断【系统生成的回答】实际主要使用的语言（必须精确区分：简体中文、繁体中文、English 等）。
   - [记录规则]：如果回答中存在语言混杂现象（如中英文混杂、繁体与简体中文混杂等），请不要只写一种语言，必须将所有出现的语言都明确列出。格式如："简体中文+English" 或 "繁体中文+简体中文"。
   - [豁免规则]：允许包含少量专有名词、英文术语或代码片段。**特别注意：允许中文回答内部存在少量的简繁体混杂（这是大模型的常见现象，不计入语种切换）。但只要出现明显的简繁体交杂使用、句子级别的中英切换，或大段异常切换，均必须严格记录为多语言混杂。**
3. language_status: 严格按以下逻辑判定：
   - "Pass"：当且仅当 output_language 与 input_language **完全一致**（如都是繁体中文，或都是简体中文），且没有发生大段无理由的跨语种突变。
   - "Failed"：出现以下任意一种情况即为 Failed：
     1. 宏观语种不一致（例如问题是繁体中文，回答却通篇是简体中文；或问题是中文，回答是英文）。
     2. 回答中途发生大面积、无理由的跨语种切换。

任务二：质量评价（务必精简，总字数严格控制在 50 到 100 字以内）
1. 准确性：回答是否准确且完整地回答了用户的问题？
2. 忠实度：回答是否严格基于【原始文档内容】，有无捏造（幻觉）或遗漏重要信息？
3. 综合评分：给出 1-10 分的评分，并给出简短的改进建议。

任务三：综合评级 (Tester Expectation)
根据回答的整体质量，给出一个 5 级评定，必须严格输出以下五个词之一：
[Excellent, Good, Pass, Poor, Failed]
（注意：如果 language_status 为 Failed，此处的评级最高只能是 Poor 或 Failed）

任务四：引用格式检测
1. Reference Link: 检查回答中是否明确提供了参考文献、来源链接或出处列表。如果有，输出 "Pass"；如果没有，输出 "N/A"。
2. Document Contain[1][2][3]: 检查回答的正文部分是否包含了类似 [1], [2], [3] 这样的数字文献引用标记。如果有，输出 "Pass"；如果没有，输出 "None"。

【请务必严格按以下 JSON 格式输出，不要包含任何 markdown 代码块标记(如```json)或其他多余文本】：
{{
    "tester_expectation": "Excellent/Good/Pass/Poor/Failed",
    "input_language": "...",
    "output_language": "...",
    "language_status": "Pass/Failed",
    "evaluation": "质量评价的详细内容...",
    "reference_link": "Pass/N/A",
    "document_contain_citations": "Pass/None"
}}

【用户问题】：
{question_text}

【目标语言】（来自Excel指定）：
{target_language}

【原始文档内容】（部分）：
{file_content[:50000]}

【系统生成的回答】：
{answer_text}
"""
    raw_result = ""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            ds_response = deepseek_client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个严谨的AI审核助手，只输出JSON格式的数据。"},
                    {"role": "user", "content": ds_prompt}
                ]
            )
            raw_result = ds_response.choices[0].message.content
            match = re.search(r'\{.*\}', raw_result, re.DOTALL)
            if match:
                clean_result = match.group(0)
                return json.loads(clean_result)
            else:
                raise ValueError("未能找到 JSON 结构")
        except Exception as e:
            safe_eval_print(f"⚠️ 第 {attempt + 1} 次调用 DeepSeek 失败: {e}")
            if attempt < max_retries - 1:
                # 👇 核心优化：增加基础等待时间，并加入随机抖动打散并发请求 👇
                base_wait = 2 ** (attempt + 1)  # 第一次等 2秒，第二次等 4秒
                jitter = random.uniform(0.5, 2.0)  # 加上 0.5 到 2 秒的随机误差
                wait_time = base_wait + jitter
                safe_eval_print(f"   ⏳ 触发防并发限流机制，等待 {wait_time:.2f} 秒后重试...")
                time.sleep(wait_time)
            else:
                # 如果解析失败，返回带有错误信息的默认字典结构
                return {
                "tester_expectation": "Parse Error(Failed)",
                "input_language": "Parse Error",
                "output_language": "Parse Error",
                "language_status": "Failed",
                "evaluation": f"API调用失败。错误: {str(e)}。 接口返回: {raw_result if raw_result else '无返回数据'}",
                "reference_link": "N/A",
                "document_contain_citations": "None"
                }
