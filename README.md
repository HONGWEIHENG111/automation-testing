## 🤖 AI Agent Web Automation & Evaluation Framework
这是一个基于 Selenium 和 DeepSeek API 构建的自动化测试框架。专门用于批量测试、抓取https://customs-demo.poffices.ai/和https://poffices.ai/ 这两个网页的生成结果（该网页是某公司做的大模型），并通过 LLM（大型语言模型）对回答质量、多语言一致性及文档引用准确性进行自动化审核与数据统计。

## ✨ 核心功能亮点
🚀 多维度自动化执行：

单线模式 (main.py)：稳定按序执行测试用例。

多线程高并发 (main_multipage.py)：使用 ThreadPoolExecutor 开启多个浏览器实例，大幅缩短大批量用例的测试时间。

多站点并行 (main_multiweb.py)：使用 threading 同时对多个不同环境（如 Demo 站与正式站）进行 A/B 测试。

🧠 深度 LLM 质量审查 (DeepSeek 驱动)：

自动比对用户输入语言与 AI 输出语言，精准侦测中英/简繁混杂现象。

基于原始文档内容，评估 AI 回答的“准确度”与“忠实度”（是否产生幻觉）。

自动检测文档引用格式（如 [1], [2] 等）。

📊 自动化数据汇总与报告：

实时将生成结果、耗时（Preparation/Completion Time）及评价写入 Excel。

测试结束后，利用 Pandas 与大模型动态提取错误样本，生成结构化的 CSV 统计报告（包含失败原因聚类与修复建议）。

## 🛡️ 强大的异常接管机制：

内置全局热键监听：按下 ESC 即可紧急刹车，安全释放浏览器控制权。

智能弹窗清理：自动跳过导览（Welcome Tour）和系统更新弹窗。

元素动态等待与重试机制，防止因网络波动导致的脚本崩溃。

## ⚙️ 环境依赖与安装
1. 基础环境
Python: >= 3.8

浏览器: 推荐安装最新版 Google Chrome（备用支持 Microsoft Edge）

2. 安装 Python 依赖包
请在终端中运行以下命令安装所需依赖：

```Bash
pip install selenium webdriver-manager PyPDF2 python-docx openpyxl pandas openai python-dotenv keybo
```
或者直接
```
pip install -r requirements.txt
```
## 🚀 使用指南
步骤 1：准备测试数据
准备一份 Excel 测试用例表（例如 Testcase_20260520_1.xlsx）。

表头必须包含：Request（问题提示词）, Selected Agent（目标代理名称）, filename（文件名）, Selected Language（目标语言，如“繁中”或“ENG”）。

如果用例包含文件上传，请将对应文件放入项目根目录的 test/ 文件夹中。

步骤 2：修改文件路径
在运行脚本前，请打开你要运行的 main*.py 文件，找到路径配置区，根据实际情况修改 Excel 文件的绝对路径：

Python
input_excel_path = r""
步骤 3：启动自动化测试
根据你的需求，在终端运行对应的脚本：

常规运行：

Bash
python main.py
多线程加速运行（适合大批量测试）：

Bash
python main_multipage.py
跨站点环境对比测试：

Bash
python main_multiweb.py
步骤 4：紧急中止
在脚本运行的任何阶段，如果需要强制停止且不希望留下残留的浏览器幽灵进程，请按下键盘上的 ESC 键，脚本将安全退出并保存已完成的进度。

📝 隐蔽水印声明
生成的 Excel 报告底层元数据已植入防伪信息，以保护代码作者知识产权。
Code Author: Henry HONG

本人学生，由于没有规划好，参与到学校的暑期实习活动，并且以实习换取毕业学分，但是该公司实在不行，各种槽点，做的项目也像骗政府钱的一样。。。
