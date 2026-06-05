import pandas as pd
import base64
import os
from openai import OpenAI
import re
import csv
import io
from datetime import datetime
from config import DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, USERNAME, PASSWORD, HOME_URL

def generate_summary_csv(excel_path, output_csv_path):
    if not os.path.exists(excel_path):
        print(f"❌ 找不到 Excel 檔案: {excel_path}")
        return

    print("📊 正在讀取測試數據並計算統計指標...")
    df = pd.read_excel(excel_path)
    total_tests = len(df)
    if total_tests == 0:
        print("⚠️ Excel 中沒有數據！")
        return

    # 1. Overall Test Result Summary (只要 Language Failed 或 Tester 覺得 Failed/Poor 就判定為 Failed)
    timeout_mask = df.get('timeout_States', pd.Series(['No'] * len(df))).astype(str).str.contains('yes', case=False,
                                                                                                  na=False)
    valid_total_tests = total_tests - timeout_mask.sum()  # 非超时任务的有效总数
    failed_mask = ((df['Language Overall Status'].astype(str).str.contains('Failed', na=False, case=False)) | \
                   (df['Tester Expectation'].astype(str).isin(['Poor', 'Failed']))) & (~timeout_mask)
    failed_count = failed_mask.sum()
    pass_count = valid_total_tests - failed_count

    # 2. 收集所有的 DeepSeek 评价与回答，用于动态生成 Failure Tags
    eval_data = []
    for idx, row in df.iterrows():
        if timeout_mask[idx]:  # 如果是超时行，直接跳过
            continue
        eval_text = str(row.get('DeepSeek评价内容', '')).strip()
        ans_text = str(row.get('answer', '')).strip()
        # 排除掉无效评价，只把有实质评价的内容喂给大模型
        if eval_text and eval_text not in ['Unknown', 'nan']:
            # 截短回答以节省 token
            short_ans = ans_text[:150].replace('\n', ' ')
            eval_data.append(f"[Test ID: {idx + 1}] Eval: {eval_text} | Ans: {short_ans}")

    all_evals_text = "\n".join(eval_data)

    # 3. Language Testing Summary
    lang_fail_count = (df['Language Overall Status'].astype(str).str.contains('Failed', na=False, case=False) & (
        ~timeout_mask)).sum()
    lang_pass_count = valid_total_tests - lang_fail_count

    # 5. Performance Summary
    comp_times = pd.to_numeric(df['Completion Time'].astype(str).str.replace('s', '', regex=False), errors='coerce')
    fastest = comp_times.min() if not comp_times.isna().all() else 0
    slowest = comp_times.max() if not comp_times.isna().all() else 0
    avg_time = comp_times.mean() if not comp_times.isna().all() else 0

    # 計算 Reference Testing 數據 ======
    if 'Reference Link' in df.columns:
        ref_pass_count = df['Reference Link'].astype(str).str.contains('Pass', case=False, na=False).sum()
        ref_fail_count = valid_total_tests - ref_pass_count
    else:
        ref_pass_count = 0
        ref_fail_count = 0

    if 'Document Contain[1][2][3]' in df.columns:
        doc_pass_count = df['Document Contain[1][2][3]'].astype(str).str.contains('Pass', case=False, na=False).sum()
        doc_none_count = valid_total_tests - doc_pass_count
    else:
        doc_pass_count = 0
        doc_none_count = 0

    # ====== 6. 呼叫 DeepSeek 生成主觀分析 ======
    print("🤖 正在呼叫 DeepSeek 提取 Example 與生成關鍵觀察 ...")
    client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)

    # 提取語言錯誤的回答樣本 (新增傳入 Test ID)
    lang_failed_mask = df['Language Overall Status'].astype(str).str.contains('Failed', case=False, na=False)
    timeout_mask = ~df.get('timeout_States', pd.Series(['No'] * len(df))).astype(str).str.contains('yes', case=False,
                                                                                                   na=False)

    lang_failed_samples = df[lang_failed_mask & timeout_mask][
        ['Input Language', 'Output Language', 'answer']].dropna().head(50)
    lang_failed_text = "\n".join(
        [
            f"[Test ID: {idx + 1}] Input: {row['Input Language']}, Output: {row['Output Language']} | Answer: {str(row['answer'])[:150]}"
            for idx, row in lang_failed_samples.iterrows()])

    # 將 tags 的統計格式化，保證 LLM 輸出準確的數量，避免幻覺
    stats_text = f"""
            總測試數: {total_tests}, 成功: {pass_count}, 失敗: {failed_count}。

            【以下是所有測試的 DeepSeek 評價內容與對應的回答樣本 (用於動態聚類 Failure Reason)】：
            {all_evals_text}

            【以下是部分語言錯誤的回答樣本，供你提取 Language Problem Example】：
            {lang_failed_text}
            """

    prompt = f"""
            我們剛剛完成了一輪 AI 助手的問答測試。以下是測試數據的統計和真實的回答樣本：
            {stats_text}

            請根據這些真實數據，幫我生成報告所需的幾個特定模塊。
            【嚴格要求】：
            1. 必須以純文本的 CSV 格式輸出，絕對不要使用 markdown 代碼塊。
            2. 語言請使用英文 (English)。
            3. Example 必須從我提供的樣本中提取或高度概括（控制在5-15個單詞以內），不要憑空捏造。
            4. 必須嚴格按照以下順序和格式輸出，並使用 ---BLOCK_START--- 和 ---BLOCK_END--- 包裹：
            5. 【重要】絕對不要自己計算 Count 和 Percentage！請提供具體的 Test ID 列表（格式嚴格為數字和逗號，例如："1, 4, 5"），後續由 Python 精確計算。

            ---BLOCK1_START---
            Failure Reason Breakdown,,,
            "One test may contain multiple issues, so percentages below are based on total test cases, not mutually exclusive.",,,
            Failure / Issue Type,Matching Test IDs,Example
            (請仔細閱讀上方提供的【所有測試的 DeepSeek 評價內容】，動態聚類出主要的錯誤類型,10个左右，最少8个，將對應的 Test ID 填入 Matching Test IDs 欄位，並從樣本中提取歸納一個 Example),,,
            ---BLOCK1_END---

            ---BLOCK2_START---
            Language Issue Breakdown,,,
        Language Problem Type,Matching Test IDs,Example,
        (請根據【語言錯誤的回答樣本】，總結出語言問題的分類，例如繁簡混雜、中英混雜等),,,
        ,,,
        Common Language Problems,,,
        Problem,Example,,
        (具體的語言現象),(對應上述問題，從樣本中提取的具體例子),,
            ---BLOCK2_END---

            ---BLOCK3_START---
            Key Observations,,,
            Area,Observation,,
            Language consistency,(根據數據寫一句總結),,
            Reference system,(根據數據寫一句總結),,
            Suggested Priority Fixes,,,
            Priority,Suggested Improvement,,
            High,(高優先級修復建議),,
            Medium,(中優先級修復建議),,
            Low,(低優先級修復建議),,
            ---BLOCK3_END---
            """

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7
        )
        full_response = response.choices[0].message.content.strip()

        # 解析三個模塊
        block1_match = re.search(r'---BLOCK1_START---\n(.*?)\n---BLOCK1_END---', full_response, re.DOTALL)
        block2_match = re.search(r'---BLOCK2_START---\n(.*?)\n---BLOCK2_END---', full_response, re.DOTALL)
        block3_match = re.search(r'---BLOCK3_START---\n(.*?)\n---BLOCK3_END---', full_response, re.DOTALL)

        # ================= 從這裡開始是替換的內容 (新增計算函數) =================

        def process_block_with_python_math(block_text, total):
            output = io.StringIO()
            writer = csv.writer(output, lineterminator='\n')
            reader = csv.reader(io.StringIO(block_text.strip()))

            is_id_mode = False

            for row in reader:
                if not row or all(not str(cell).strip() for cell in row):
                    continue
                # 替換表頭
                if "Matching Test IDs" in row:
                    is_id_mode = True
                    new_row = []
                    for cell in row:
                        if cell.strip() == "Matching Test IDs":
                            # 展開成三列，表頭統一使用你要求的名稱
                            new_row.extend(["Count", "Percentage", "Label"])
                        else:
                            new_row.append(cell)
                    writer.writerow(new_row)
                    continue

                # 攔截並計算數字（智能合併被 CSV 誤拆分的 ID）
                if is_id_mode and len(row) > 1:
                    extracted_ids = []
                    example_texts = []

                    # 遍歷從第二列開始的所有格子
                    for cell in row[1:]:
                        clean_cell = cell.strip()
                        # 如果這格只有數字、逗號或空格，說明它是被誤拆分的 ID
                        if clean_cell and re.match(r'^[\d\s",]+$', clean_cell):
                            # 把抓到的數字塞進 ID 列表
                            extracted_ids.extend(re.findall(r'\d+', clean_cell))
                        else:
                            # 只要遇到包含字母/中文的格子，就當作 Example 等後續文本
                            example_texts.append(cell)

                    # 只要有抓到任何數字，就執行重新計算與合併
                    if extracted_ids:
                        # 轉為整數、去重並排序，確保 ID 乾淨且不重複
                        unique_ids = sorted(list(set(map(int, extracted_ids))))
                        count = len(unique_ids)
                        percentage = f"{(count / total):.2%}" if total > 0 else "0.00%"
                        ids_str = ", ".join(map(str, unique_ids))  # 將數字重新用逗號拼接

                        # 重新組裝這行：[問題類型, 個數, 占比, 合併後的標籤] + [剩下的文本]
                        new_row = [row[0], count, percentage, ids_str] + example_texts
                        writer.writerow(new_row)
                        continue

                # 其他無關行原樣輸出
                writer.writerow(row)
            return output.getvalue().strip()

        # 套用新的計算函數
        failure_reason_csv = process_block_with_python_math(block1_match.group(1),
                                                            total_tests) if block1_match else "Failure Reason Breakdown,,,\nError,Parsing Failed,,\n"
        common_lang_csv = process_block_with_python_math(block2_match.group(1),
                                                         total_tests) if block2_match else "Common Language Problems,,,\nProblem,Example,,\nError,Parsing Failed,,\n"
        observations_csv = block3_match.group(
            1).strip() if block3_match else "Key Observations,,,\nArea,Observation,,\nError,Parsing Failed,,\n"
        # ================= 替換到這裡結束 =================

    except Exception as e:
        # (這裡及以下的報錯處理保留不動)
        print(f"⚠️ DeepSeek 生成總結失敗: {e}")
        failure_reason_csv = "Failure Reason Breakdown,,,\nError,API Failed,,\n"
        common_lang_csv = "Common Language Problems,,,\nProblem,Example,,\nError,API Failed,,\n"
        observations_csv = "Key Observations,,,\nArea,Observation,,\nError,API Failed,,\n"
    # ====== 7. 拼接最終的 CSV 內容並寫入檔案 ======
    print("📝 正在生成最終的 CSV 報告...")
    current_time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    meta_info_lines = [
        "System Login & Generation Info,,,",
        f"Login Website,{HOME_URL},,",
        f"Login Account,{USERNAME},,",
        f"Login Password,{PASSWORD},,",
        f"Generation Time,{current_time_str},,",
        ",,,"  # 空行分隔线
    ]
    # 1. 總結部分
    summary_lines = [
        "Overall Test Result Summary,,,",
        "Item,Count,Percentage,",
        f"Total Tests,{total_tests},100%,",
        f"Pass,{pass_count},{pass_count / total_tests:.2%} ",
        f"Failed,{failed_count},{failed_count / total_tests:.2%} ",
        ",,,"
    ]
    valid_div = valid_total_tests if valid_total_tests > 0 else 1
    # 2. 語言測試部分
    lang_lines = [
        "Language Testing Summary,,,",
        "Overall Language Status,,,",
        "Status,Count,Percentage,",

        f"Pass,{lang_pass_count},{lang_pass_count / valid_div:.2%} ",
        f"Failed,{lang_fail_count},{lang_fail_count / valid_div:.2%} ",
        ",,,",
    ]

    # ======  Reference 測試部分 ======
    ref_lines = [
        "Reference Testing Summary,,,",
        "Reference Output Status,,,",
        "Status,Count,Percentage,",
        f"Pass,{ref_pass_count},{ref_pass_count / valid_div:.2%} ",
        f"Failed,{ref_fail_count},{ref_fail_count / valid_div:.2%} ",
        ",,,",
        "Document Citation Status,,,",
        "Status,Count,Percentage,",
        f"Pass,{doc_pass_count},{doc_pass_count / valid_div:.2%} ",
        f"None,{doc_none_count},{doc_none_count / valid_div:.2%} ",
        ",,,"
    ]

    # ====== 3. 性能部分  ======
    perf_lines = [
        "Performance Summary,,,",
        "Metric,Result,,",
        f"Fastest Response,{fastest}s,,",
        f"Slowest Response,{slowest}s,,",
        f"Average Loading Time,~{avg_time:.1f}s,,",
        ",,,"
    ]
    # ================= Timeout 測試部分 =================
    # ================= Timeout 測試部分 =================
    if 'timeout_States' in df.columns:
        # 1. 宏观超时判定 (只要有 yes 就是超时)
        timeout_mask = df['timeout_States'].astype(str).str.contains('yes', case=False, na=False)
        timeout_count = timeout_mask.sum()
        no_timeout_count = total_tests - timeout_count

        timeout_ids = df[timeout_mask].index + 1
        timeout_ids_str = ", ".join(map(str, timeout_ids.tolist())) if timeout_count > 0 else "None"

        no_timeout_ids = df[~timeout_mask].index + 1
        no_timeout_ids_str = ", ".join(map(str, no_timeout_ids.tolist())) if no_timeout_count > 0 else "None"

        # 2. 细分：总时间超时的数量与号数
        total_timeout_mask = df['timeout_States'].astype(str).str.contains('总时间超时', na=False)
        total_time_out = total_timeout_mask.sum()
        total_timeout_ids = df[total_timeout_mask].index + 1
        total_timeout_ids_str = ", ".join(map(str, total_timeout_ids.tolist())) if total_time_out > 0 else "None"

        # 3. 细分：生成超时的数量与号数
        gen_timeout_mask = df['timeout_States'].astype(str).str.contains('生成超时', na=False)
        gen_time_out = gen_timeout_mask.sum()
        gen_timeout_ids = df[gen_timeout_mask].index + 1
        gen_timeout_ids_str = ", ".join(map(str, gen_timeout_ids.tolist())) if gen_time_out > 0 else "None"

    else:
        timeout_count = 0
        no_timeout_count = total_tests
        total_time_out = 0
        gen_time_out = 0
        timeout_ids_str = "None"
        no_timeout_ids_str = ", ".join(map(str, range(1, total_tests + 1))) if total_tests > 0 else "None"
        total_timeout_ids_str = "None"
        gen_timeout_ids_str = "None"

    # 4. 写入 CSV 行
    timeout_lines = [
        "Timeout Summary,,,",
        "Status,Count,Percentage,Matching Test IDs",
        f"No Timeout,{no_timeout_count},{no_timeout_count / total_tests:.2%} ,\"{no_timeout_ids_str}\"",
        f"Has Timeout (Total),{timeout_count},{timeout_count / total_tests:.2%} ,\"{timeout_ids_str}\"",
        f" -> Total Time Exceeded (600s),{total_time_out},N/A,\"{total_timeout_ids_str}\"",  # 👈 新增号数输出
        f" -> Generation Exceeded (120s),{gen_time_out},N/A,\"{gen_timeout_ids_str}\"",  # 👈 新增号数输出
        ",,,"
    ]
    # ====== 按照要求的順序精準拼接 ======
    final_csv_content = (
            "\n".join(meta_info_lines) + "\n" +
            "\n".join(summary_lines) + "\n" +
            failure_reason_csv + "\n,,,\n" +
            "\n".join(lang_lines) + "\n" +
            common_lang_csv + "\n,,,\n" +
            "\n".join(ref_lines) + "\n" +  # 插入 Reference 模塊
            "\n".join(perf_lines) + "\n" +  # 插入 性能 模塊
            "\n".join(timeout_lines) + "\n" +  # <--- 新增：插入 Timeout 模塊
            observations_csv
    )

    with open(output_csv_path, 'w', encoding='utf-8-sig') as f:
        f.write(final_csv_content)

    print(f"✅ 大功告成！匯總報告已保存至: {output_csv_path}")


# 允許腳本獨立運行測試
if __name__ == "__main__":
    # 假設這兩個檔案與腳本在同一目錄下
    test_in = "evaluation_results.xlsx"
    test_out = "Summary_Test.csv"
    generate_summary_csv(test_in, test_out)
