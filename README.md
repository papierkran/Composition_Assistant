# OCR 图像文字识别工具

本项目是一个基于 Python 的图形化工具，支持使用 [讯飞开放平台](https://www.xfyun.cn/services/ocr_general) 的 **手写文字识别 API** 对指定文件夹下的图片进行批量 OCR 识别，并自动生成 Word 文档。

支持自定义 API 参数配置，自动保存配置，下次启动自动填充。图形化界面友好，操作简便，适合教师、教辅、资料归档等使用场景。

支持 deepseek api 的错别字改正

即将更新多工具选择进行ocr（画饼）

---

## 📦 功能说明

- 支持识别 `.jpg`, `.png`, `.bmp`, `.jpeg` 格式图片
- 自动识别当前文件夹或其子文件夹中的图片内容
- 图片识别内容自动写入 Word（`.docx`），字体：宋体，小四，段前段后为 0，最小行距 12 磅
- 每张图片自动标注“修改前”、“修改后”结构，并分页
- 图片所在文件夹名写入 Word 中作为“——姓名”标注
- 支持 GUI 操作，支持打包为 `.exe` 使用，无需 Python 环境
- (V0.3更新) 简易自然段逻辑分割
- (V0.5更新) 支持deepseek api 的错别字改正 (注意config.json的配置文件进行配置)

---

## 🖼️ 图形界面功能

- 输入讯飞 `APPID`, `API_KEY`, `OCR接口URL`
- 选择或手动输入图片所在目录路径
- 点击“开始识别”按钮自动批量处理并生成 Word 文件
- 所有配置自动保存，下次启动自动加载

---

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```
或手动安装：
```
pip install PyQt5 python-docx requests
```

### 2. 运行程序
```
python ocr_gui.py
```

## 📝 配置说明
程序会自动在当前目录生成一个配置文件：
```json
config.json
```

## 📄 输出格式说明

每张图片对应一个“修改前：”和“修改后：”段落结构

段落之间添加分页符

在“修改前：”上方插入 ——姓名（文件夹名） 居中段落

所有图片识别结果拼接在同一个 Word 中，文件名为对应文件夹名称

## 🔐 目前 ocr API 获取方式
请在讯飞开放平台申请并获取：

APPID

API_KEY

OCR 接口地址（默认可用）

申请地址：https://www.xfyun.cn/services/ocr_general

## 📌 注意事项

仅支持中文/英文手写识别（建议图片清晰，避免旋转）

单张图识别失败不会影响整体执行，会跳过

若图片较大或数量过多，识别需等待数秒

单张图片不能超过5MB

## TODO

段落识别

自动改错别字

接入更多api

本地运行ocr



## 📃 License
本项目仅供学习和教育用途，禁止商业或非法用途。API 使用请遵守讯飞开发者服务协议。

---

## 🔄 PySide6 重构版 (V2.0)

本项目已从 CustomTkinter 迁移到 PySide6 (Qt)，提供更好的性能和高DPI支持。

### 主要改进
- **性能提升**：原生 Qt 渲染，窗口缩放流畅，4K屏不再卡顿
- **高DPI支持**：自动适配高分辨率显示器
- **更稳定的布局**：使用 Qt 的 QVBoxLayout/QHBoxLayout，布局更可靠
- **更好的滚动性能**：QScrollArea 替代自定义滚动框架

### 文件结构
```
Composition_OCR_Assistant/
├── ocr_gui.py              # PySide6 主程序
├── config_editor_ui.py     # PySide6 配置编辑器
├── ocr_main.py             # 核心处理逻辑
├── config_migrate.py       # 配置迁移工具
├── llm_client.py           # LLM 客户端
├── config.json             # 配置文件
├── app.ico                 # 应用图标
├── ocr_gui.spec            # PyInstaller 打包配置
├── dist/
│   └── ocr_gui.exe         # 打包好的可执行文件
└── 旧版CTK/                # 旧版 CustomTkinter 代码备份
    ├── ocr_gui_ctk.py
    ├── config_editor_ui_ctk.py
    └── ...
```

### 运行方式
```bash
python ocr_gui.py
```

### 打包为 exe
```bash
pyinstaller --clean ocr_gui.spec
```
生成的 exe 位于 `dist/ocr_gui.exe`

### 依赖
```
PySide6>=6.5.0
python-docx
openai
Pillow
```

### 功能特性
- **双页面切换**：图片转作文 / docx作文处理
- **可折叠配置区**：百度图片矫正、OCR配置可收起展开
- **任务队列管理**：支持批量添加、删除、刷新任务
- **AI 错别字修正**：集成 DeepSeek/OpenAI 等 LLM API
- **AI 作文修改**：支持字数控制和自定义提示词
- **多 Provider 支持**：可添加自定义 AI 服务提供商