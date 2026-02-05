# Copilot / AI Agent 指南 — Composition_OCR_Assistant

目标：帮助 AI 迅速理解仓库结构、关键约定、运行与构建流程，以及小心修改时容易破坏的隐式契约。

- **大体架构**：
  - `ocr_gui.py`：图形化前端（CustomTkinter + tkinter），负责读取/保存配置并在独立线程中调用 `ocr_main.process_all`。
  - `ocr_main.py`：核心处理流程：调用讯飞 OCR（HTTP）、合并文本段落、生成 Word（`python-docx`），并按配置可调用 DeepSeek/OpenAI 风格的 API 进行纠错与二次编辑。
  - `paddle处理.py`：基于 `PaddleOCR` 的替代实现，保留类似的段落合并与 Word 生成逻辑，供离线/本地模型替代线上讯飞服务。
  - `config.json`：默认存放运行时配置（OCR/DEEPSEEK/APP），程序会读取此文件的键来决定行为。

- **关键文件与位置（快速定位）**
  - `ocr_gui.py` — GUI 和用户交互（启动入口：`python ocr_gui.py`）
  - `ocr_main.py` — 后端 OCR/文档处理逻辑（可直接运行或由 GUI 调用）
  - `paddle处理.py` — 本地 PaddleOCR 版本
  - `ocr_gui.spec`, `ocr_main.spec` — PyInstaller 打包配置（仓库包含 .spec，可用来复现 `.exe`）

- **重要运行 / 构建命令**
  - 运行 GUI：`python ocr_gui.py`（推荐，用于交互式调试）
  - 运行无 GUI：`python ocr_main.py`（会提示输入要处理的目录）
  - 打包（已有 .spec）：
    - `pyinstaller ocr_gui.spec`
    - `pyinstaller ocr_main.spec`
  - 依赖安装（根据 README 与代码，主要依赖）：`pip install requests python-docx customtkinter openai paddleocr`（如项目含 `requirements.txt`，优先使用）

- **配置与常见坑**
  - 代码中有硬编码的 `CONFIG_FILE` 路径（例如 `D:\person_data\ocer助手\presson.json`），并且在多处被注释替换为 `config.json`。修改或测试时务必：
    - 优先使用仓库根目录的 `config.json`，或把 `CONFIG_FILE` 常量改为相对路径 `config.json`。
    - `ocr_gui.py` 与 `ocr_main.py` 两处都需要保持一致（否则 GUI 与后端可能读不同的配置）。
  - 必填配置项（`config.json` 中）：`OCR.URL`, `OCR.APPID`, `OCR.API_KEY`, `APP.ROOT_DIR`。若启用 DeepSeek，需填 `DEEPSEEK.API_KEY`。
  - UI 与存储：UI 会把 API Key 隐藏到配置文件或内存（`hidden_api_keys`），因此修改密钥逻辑时注意 `entries_map`/`hidden_api_keys` 的交互。

- **约定与契约（修改时务必小心）**
  - Word 文档结构依赖精确字符串：`修改前：` 与 `修改后：`。多处函数（`extract_before_text`, `clear_before_text`, `insert_before_text` 等）以此定位写入/替换位置。不要轻易改动这些文本或文档流结构，除非同时更新所有文档操作函数。
  - 段落合并算法（OCR → 段落）在 `ocr_main.ocr_and_extract_text` 与 `paddle处理.ocr_and_extract_text` 中实现：短行合并、以句末标点和长度判断分段。任何改进须在两处保持一致以避免格式差异。
  - DeepSeek / 第二步编辑使用 `openai.OpenAI` 客户端並通过 `base_url` 覆盖实现。接口返回的文本被按行 split 成段落后插入文档，需对 AI 返回值做防护（例如过滤掉包含“修改前：”等特殊关键词）。

- **调试提示**
  - 若出现“未找到 config.json 或相关配置文件”的错误，先检查 `CONFIG_FILE` 常量是否指向有效路径。
  - 日志回调在 GUI 中通过 `append_log`，在后端使用 `log_callback` 参数（默认是 `print`）。调试时可传自定义 callback 捕获日志。

- **常见改动示例**
  - 将运行时配置改为相对路径：在 `ocr_gui.py` 和 `ocr_main.py` 将 `CONFIG_FILE = "D:\\person_data\\ocer助手\\presson.json"` 改为 `CONFIG_FILE = "config.json"`，并确保 `config.json` 存在于仓库根目录。
  - 切换为 PaddleOCR 本地处理：在测试分支里，参考 `paddle处理.py` 的 `ocr_and_extract_text` 与 `process_all`，或在 GUI 中添加一个选择项来切换不同后端实现。

- **不要从代码以外推断的事**
  - README 中提到 `PyQt5`，但当前 GUI 使用 `customtkinter`/`tkinter`；以实际源码为准。
  - 不要假设 `requirements.txt` 完整（请以导入包为准，并在必要时运行 `pip freeze` 或手动安装缺失包）。

- 若有需要，我可以：
  - 将 README 中关于 PyQt5 的说明修正为使用 `customtkinter`。
  - 在代码中统一 `CONFIG_FILE` 为相对 `config.json` 并更新相关读取/保存逻辑（我可以提交补丁）。

请指出是否有希望补充的运行环境、CI、或本地测试命令。
