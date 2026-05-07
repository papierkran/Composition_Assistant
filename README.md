# Composition Assistant 作文修改助手

一个面向作文图片批量处理的 PySide6 图形化工具。它可以扫描作文文件夹中的图片，调用讯飞手写 OCR 识别文字，生成 Word 文档，并按需接入 OpenAI-compatible API 做错别字修正和作文改写。

适合教师、教辅、资料整理等批量处理场景。

---

## 主要功能

- 图片批量 OCR：支持 `.jpg`、`.jpeg`、`.png`、`.bmp`
- 自动扫描任务：可识别当前作文文件夹，或其一级子文件夹中的图片任务
- Word 输出：按作文文件夹名生成同名 `.docx`
- AI 错别字修正：支持 DeepSeek、OpenAI 或自定义 OpenAI-compatible Provider
- AI 作文改写：支持自定义 Prompt，并可设置目标字数范围
- 百度图片矫正：可在 OCR 前自动矫正倾斜/弯曲文档图片
- 任务队列：显示学生姓名、文件路径、作文名称、修改前字数、当前步骤、修改后字数、状态、实时日志
- 多线程处理：默认同时处理 3 个任务，队列日志按任务行独立更新
- 失败重试：只有“已完成”的任务会跳过，“失败”的任务再次开始时会重新处理
- 手动重新加入：选中任务后点击“重新加入”，可把任务改回等待处理
- 配置编辑器：可在 GUI 中管理 OCR、AI Provider、Prompt 等配置

---

## 界面说明

### 图片转作文

这是主处理页，用于从作文图片生成 Word，并可继续执行 AI 修正/改写。

常用流程：

1. 填写或选择“作文文件夹”路径。
2. 按需展开“百度图片矫正”“OCR 识别配置”“AI 错别字修正”“AI 修改作文”。
3. 点击“开始处理（自动读取路径下任务并开始）”。
4. 查看任务队列中的“当前步骤”“状态”和“实时日志”。

任务队列按钮：

- `添加`：手动选择一个含图片的任务文件夹加入队列。
- `删除`：删除选中的队列项，并清除其完成/失败标记。
- `重新加入`：把选中的任务重新改为待处理；正在处理中的任务会跳过。
- `读取`：扫描当前作文文件夹，把新任务加入队列。
- `刷新队列`：清理无效路径并重绘队列。

开始处理规则：

- 点击开始时会重新扫描当前作文文件夹。
- 已在队列中的任务不会重复添加。
- 后续新增任务会追加到队列。
- 如果已有任务正在处理，不会启动第二个调度器；新任务会排队等待线程池处理。
- 状态为“已完成”的任务会跳过。
- 状态为“失败”或“待完成”的任务会重新进入处理队列。

### docx作文处理

用于对已有 Word/图片文件执行文档后处理流程。

可选步骤包括：

- DOC 转 DOCX
- 清除空格
- AI 改作文
- 添加标签
- 格式化
- 修改作者

处理日志可折叠展开，适合单独对已有文档做二次整理。

---

## 输出格式

图片转作文会在任务文件夹中生成同名 Word 文件：

```text
任务文件夹/
├── 1.jpg
├── 2.jpg
└── 任务文件夹名.docx
```

Word 内容结构：

- 文件夹名会作为姓名/标题标注写入文档。
- OCR 内容写入“修改前：”区域。
- 文档预留“修改后：”区域。
- 启用 AI 错别字修正后，会把修正内容插入到对应区域。
- 启用 AI 作文改写后，会把改写结果写入“修改后：”区域。

---

## 配置说明

配置加载优先级：

1. `D:\person_data\ocer助手\presson.json`
2. 程序所在目录的 `config.json`
3. 默认配置

GUI 保存配置时会优先保存到个人配置目录；如果个人配置目录不存在，则保存到程序所在目录的 `config.json`。

主要配置项：

- `OCR.XFYUN.URL`：讯飞 OCR 接口地址
- `OCR.XFYUN.APPID`：讯飞 APPID
- `OCR.XFYUN.API_KEY`：讯飞 API Key
- `OCR.BAIDU_CORRECTION`：百度图片矫正配置
- `LLM.PROVIDERS`：AI Provider 列表
- `LLM.TASKS.typo_fix`：错别字修正任务配置
- `LLM.TASKS.editor`：作文改写任务配置
- `APP.ROOT_DIR`：默认作文文件夹路径

---

## 运行方式

安装依赖：

```bash
pip install PySide6 python-docx requests openai Pillow
```

运行 GUI：

```bash
python ocr_gui.py
```

命令行处理入口：

```bash
python ocr_main.py <作文文件夹路径>
```

常用参数：

```bash
python ocr_main.py <作文文件夹路径> --config config.json
python ocr_main.py <作文文件夹路径> --no-deepseek
python ocr_main.py <作文文件夹路径> --no-editor
python ocr_main.py <作文文件夹路径> --debug
```

---

## 打包为 EXE

```bash
pyinstaller --clean ocr_gui.spec
```

生成文件位于：

```text
dist/ocr_gui.exe
```

---

## 文件结构

```text
Composition_OCR_Assistant/
├── ocr_gui.py              # PySide6 主程序
├── ocr_main.py             # OCR 与 Word/AI 处理核心逻辑
├── config_editor_ui.py     # 配置编辑器
├── config_migrate.py       # 配置迁移工具
├── llm_client.py           # OpenAI-compatible LLM 客户端
├── baidu_image_corrector.py
├── ocr_gui.spec            # PyInstaller 打包配置
├── config.json             # 本地配置文件
├── app.ico                 # 应用图标
├── dist/                   # 打包输出
└── 旧版CTK/                # 旧版 CustomTkinter 代码备份
```

---

## 注意事项

- 讯飞手写 OCR 单张图片大小不要超过接口限制，建议图片清晰、方向正确。
- 启用百度图片矫正会改变处理前的图片流程，建议先用少量样本验证效果。
- 启用 AI 错别字修正或作文改写会明显增加处理时间，并消耗对应 API 额度。
- 多线程默认并发数为 3；如果 API 限流频繁，可在 `ocr_gui.py` 中调整 `self.max_parallel_tasks`。
- 任务队列状态以 GUI 运行期间为准；重启程序后不会恢复上次队列运行状态。

---

## License

本项目仅供学习和教育用途，禁止商业或非法用途。OCR 和 AI API 使用请遵守对应服务商协议。
