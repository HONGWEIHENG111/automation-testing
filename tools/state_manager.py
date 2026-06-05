# state_manager.py
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
def handle_post_send(driver, current_state):
    """
    阶段 2：负责处理点击“发送”按钮之后的特殊状态逻辑
    """
    if current_state == "1":
        print("   🔍 [状态1专属]: 发送后，正在手动选择 General Agent...")
        # 状态 1 发送后，需要手动选择 Agent
        general_radio = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//input[@aria-label='General' or @aria-label='通用AI代理模式']"))
        )
        driver.execute_script("arguments[0].click();", general_radio)
        time.sleep(1)

        confirm_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Confirm & Generate')]"))
        )
        driver.execute_script("arguments[0].click();", confirm_btn)

    else:
        pass
def execute_state(driver, current_state, target_agent=""):
    """
    根据 current_state 执行对应的自动化操作逻辑
    """
    # state1 就是选择 Agent Finder Mode并且关闭Auto mode
    if current_state == "1":
        print("   🔍 正在选择 Agent Finder Mode...")
        mode_radio = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'mode-header') and (.//span[text()='Agent Finder Mode'] or .//span[text()='代理尋找模式'])]//div[contains(@class, 'mode-radio')]"))
        )
        driver.execute_script("arguments[0].click();", mode_radio)
        time.sleep(0.5)

        auto_mode_checkbox = driver.find_element(By.ID, "finder-auto-mode")
        if auto_mode_checkbox.is_selected():
            slider_btn = auto_mode_checkbox.find_element(By.XPATH, "./following-sibling::span[@class='slider']")
            driver.execute_script("arguments[0].click();", slider_btn)
        time.sleep(0.5)

    # state2 就是选择 Agent Finder Mode并且开启Auto mode
    elif current_state == "2":
        print("🟢 当前运行：状态 2 (Agent Finder Mode + Auto开启 + 自动生成)")
        try:
            print("   🔍 正在寻找弹窗内的 mode-radio...")
            mode_radio = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'mode-header') and (.//span[text()='Agent Finder Mode'] or .//span[text()='代理尋找模式'])]//div[contains(@class, 'mode-radio')]"))
            )
            driver.execute_script("arguments[0].click();", mode_radio)
            time.sleep(0.5)

            print("   🔍 正在处理 Auto Mode 开关...")
            auto_mode_checkbox = driver.find_element(By.ID, "finder-auto-mode")
            if not auto_mode_checkbox.is_selected():
                slider_btn = auto_mode_checkbox.find_element(By.XPATH, "./following-sibling::span[@class='slider']")
                driver.execute_script("arguments[0].click();", slider_btn)
                print("   ✅ 已将 Auto Mode 从【关闭】切换为【开启】")
            else:
                print("   ✅ Auto Mode 已经是【开启】状态，保持不变")
            time.sleep(0.5)

        except Exception as set_err:
            print(f"   ⚠️ [警告] 设置 Agent Finder 面板时超时或失败，正在跳过设置继续执行: {set_err}")
    # state3 就是选择 General Agent Mode
    elif current_state == "3":
        print("🟢 当前运行：状态 3 (General Agent Mode + 自动生成)")
        try:
            print("   🔍 步骤 B: 正在寻找并选择 General Agent Mode...")
            general_mode_radio = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'mode-header') and (.//span[contains(text(), 'General Agent Mode')] or .//span[contains(text(), '通用AI代理模式')])]//div[contains(@class, 'mode-radio')]"))
            )
            driver.execute_script("arguments[0].click();", general_mode_radio)
            time.sleep(0.5)

        except Exception as set_err:
            print(f"   ⚠️ [警告] 设置 General Agent 面板时超时或失败，正在跳过设置继续执行: {set_err}")

    # state4 就是选择  Agent Master Mode并且开启 Auto Mode
    elif current_state == "4":
        print(f"🟢 当前运行：状态 {current_state} (Agent Master Mode 通用步骤)")
        try:
            # 1. 寻找并选择 Agent Master Mode 的单选框
            print("   🔍 正在选择 Agent Master Mode...")
            master_mode_radio = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH,
                                                "//div[contains(@class, 'mode-header') and (.//span[text()='Agent Master Mode'] or .//span[text()='代理自選模式'])]//div[contains(@class, 'mode-radio')]"))
            )
            driver.execute_script("arguments[0].click();", master_mode_radio)
            time.sleep(0.5)

            # 2. 处理 Agent Master 特属的 Auto Mode 开关 (id="master-auto-mode")
            try:
                # 尝试寻找该元素，如果页面上没有该元素，会直接触发错误并跳到 except 块
                master_auto_checkbox = driver.find_element(By.ID, "master-auto-mode")

                # 如果找到了且该开关在页面上是可见的，则执行原有的点击开启逻辑
                if master_auto_checkbox.is_displayed():
                    if not master_auto_checkbox.is_selected():
                        slider_btn = master_auto_checkbox.find_element(By.XPATH,
                                                                       "./following-sibling::span[@class='slider']")
                        driver.execute_script("arguments[0].click();", slider_btn)
                        print("   ✅ 已将 Master Auto Mode 从【关闭】切换为【开启】")
                    else:
                        print("   ✅ Master Auto Mode 已经是【开启】状态，保持不变")
                    time.sleep(0.5)
            except Exception:
                # 如果没有该按键或不可见，直接打印提示，不进行任何点击，无缝进入下一步
                print("   ℹ️ 未检测到或未跳出 Auto Mode 开关，自动跳过此点击动作。")

            # 3. 点击进入 Agent Master 选项卡 (tab)
            print("   🔍 正在点击 Agent Master 侧边栏菜单...")
            agent_master_tab = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//li[@data-tab='tab-agent-master']"))
            )
            driver.execute_script("arguments[0].click();", agent_master_tab)
            time.sleep(1)  # 等待右侧面板切换完成

            print("   🧹 正在点击 Clear All 清空历史选择...")
            try:
                clear_all_btn = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "agent-master-clear-all"))
                )
                driver.execute_script("arguments[0].click();", clear_all_btn)
                time.sleep(0.5)
                print("   ✅ 历史选择已成功清空")
            except Exception as clear_err:
                print(f"   ⚠️ 未找到 Clear All 按钮或清空失败（可能当前列表本就是空的）: {clear_err}")

            print("   🔍 正在点击 Agent 搜索框...")
            search_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "agent-master-search"))
            )
            driver.execute_script("arguments[0].click();", search_input)
            time.sleep(0.5)
            # 4. 点击并输入搜索框
            print(f"   🔍 正在搜索指定的 Agent: {target_agent if target_agent else '空'}")

            if target_agent:

                search_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "agent-master-search"))
                )
                # 2. 用 JS 一步到位：清空 -> 赋值 -> 触发前端双向绑定事件
                driver.execute_script("""
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """, search_input, target_agent)

                print(f"   ✅ [JS 注入]: 已在后台成功写入搜索词 '{target_agent}'")

                # 留点时间给网页把 General 藏起来，把 Chatbot 刷出来
                time.sleep(2)

                # 5. 终极精准匹配：直接锁定 button 的 title 属性
                print(f"   ➕ 正在精准匹配并添加 Agent: {target_agent}")
                try:
                    # 先等待带有加号 class 的按钮出现在网页中
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "am-agent-add-btn"))
                    )

                    # 把当前网页里所有的加号按钮都抓出来
                    add_buttons = driver.find_elements(By.CLASS_NAME, "am-agent-add-btn")

                    clicked = False
                    # 把我们期待的名字转小写，比如 "chatbot"
                    target_lower = target_agent.strip().lower()
                    # 拼凑出我们期待的 title 完整形态，比如 "add chatbot"
                    expected_title = f"add {target_lower}"

                    for btn in add_buttons:

                        # 【核心防御 2】：获取这个按钮专属的 title 属性
                        # 比如网页上是 title="Add Chatbot"，提取出来就是 "Add Chatbot"
                        title_text = btn.get_attribute("title")

                        # 把提取出的 title 也转成小写，进行严格比对
                        if title_text and target_lower in title_text.strip().lower():
                            # 找到了！这个就是我们要的加号！
                            driver.execute_script("arguments[0].click();", btn)
                            print(f"   ✅ 已成功精准点击对应的加号按钮 (识别到属性: {title_text})")
                            clicked = True
                            break  # 点完收工，跳出循环

                    if not clicked:
                        error_msg = f"未找到对应Agent的加号按钮 (预期 title: {expected_title})"
                        print(f"   ⚠️ {error_msg}")
                        raise Exception(error_msg)

                    time.sleep(1)

                except Exception as add_err:
                    print(f"   ⚠️ 查找或点击加号按钮时发生异常: {add_err}")
                    raise add_err
            else:
                print("   ⚠️ 当前用例的 Selected Agent 为空，跳过输入。")

        except Exception as set_err:
            print(f"   ⚠️ [警告] 设置 Agent Master 面板时超时或失败: {set_err}")
    else:
        print(f"❌ 未知的状态值: {current_state}，请检查配置！")
        raise ValueError("Unknown CURRENT_STATE")
