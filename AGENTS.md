# AGENTS.md

本文件为 AI 编程助手（Agent）在参与本项目（`usdoc`）维护、开发与重构时的标准指引文档。请在修改代码前仔细通读本指南。

---

## 1. 项目定位与核心功能

**usdoc** 是一个基于 **FastAPI + Jinja2 + docxtpl + python-docx + openpyxl** 的智能化文档生成系统：
- **模板驱动渲染**：用户上传包含占位符（如 `{{project_name}}`）的 `.docx` 模板。
- **大模型驱动填充**：后端自动解析模板中包含的所有占位符变量，调用 OpenAI 兼容协议的大模型（LLM）生成严格匹配变量 Schema 的 JSON 数据。
- **文档输出与导出**：
  - 单篇 Word 生成（支持 FastAPI `BackgroundTasks` 后台异步生成与状态轮询记录）。
  - 单篇 PDF 转换（支持本地 `libreoffice` 或 `docx2pdf` 工具）。
  - Excel 批量生成（解析 `.xlsx` 行数据，结合模板和提示词批量渲染 Word，并打包成 `.zip` 下载，具备单行重试与容错回退机制）。
- **单文件便携化发布**：支持通过 PyInstaller 打包为独立的 Windows EXE 文件（`dist/usdoc.exe`），以及 GitHub Actions CI 自动化 Release。

---

## 2. 目录结构与关键文件说明

```text
.
├── main.py                     # 核心业务服务：FastAPI 路由、AI 交互、模板解析、Word 渲染、记录存储
├── run_usdoc.py                # 服务启动入口：通过 uvicorn.run 启动应用，支持 PORT 环境变量
├── requirements.txt            # Python 依赖清单
├── usdoc.spec                  # PyInstaller 单文件打包规范配置
├── templates/
│   └── index.html              # 前端单页面交互界面（Jinja2 模板）
├── static/
│   └── style.css               # 前端现代卡片风格样式表
├── .github/
│   └── workflows/
│       └── build-windows-release.yml  # GitHub Actions 自动化构建 Windows EXE 与 Release 工作流
├── .gitignore                  # Git 忽略规则（忽略虚拟环境、临时上传/输出文件、日志等）
├── uploads/                    # 运行时动态创建：用户上传的模板与临时 Excel 存放目录
├── outputs/                    # 运行时动态创建：生成的 Word、PDF、ZIP 文件存放目录
└── generation_records.json     # 运行时持久化：任务历史生成记录（JSON 文件）
```

---

## 3. 技术栈与运行依赖

- **语言环境**：Python 3.10+ / 3.11
- **Web 框架**：FastAPI + Uvicorn
- **文档与表格处理**：
  - `docxtpl`：基于 Jinja2 语法的 Word 模板渲染核心。
  - `python-docx`：Word 文档文本提取、段落及表格遍历。
  - `openpyxl`：Excel `.xlsx` 文件流式读取与解析。
- **AI 与网络请求**：`requests`（调用 OpenAI 兼容接口）。
- **数据序列化与验证**：`pydantic`（请求模型声明）、内置 `json` 与 `re`。
- **外部可选工具**：
  - `libreoffice` 或 `docx2pdf`：用于 Word 转 PDF 功能。未安装时该功能返回 500 友好提示。

---

## 4. 架构设计与核心执行流程

### 4.1 目录与打包兼容性（`resolve_base_dir`）
由于项目支持 PyInstaller 打包：
- 静态资源与模板路径：使用 `sys._MEIPASS`（打包状态下解压目录）或当前文件所在路径。
- 数据存储目录（`DATA_DIR`）：打包时必须使用当前运行目录 `Path.cwd()`，确保生成的 `uploads/`、`outputs/`、`generation_records.json` 能够正确在用户本地留存，而不是随临时目录销毁。

### 4.2 模板变量规范与校验
- **变量提取逻辑**：遍历 Word 的所有段落（`paragraphs`）与所有表格单元格（`tables`），正则匹配 `{{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}`。
- **非法占位符拦截**：检测非标准变量格式（如包含特殊字符、非法命名），在渲染前抛出明确的 HTTP 400 错误提示，防止模板渲染引擎静默崩溃。

### 4.3 AI 请求与结构化 JSON 约束
1. 后端根据模板中检测到的 `variables` 列表，动态构造严格的 JSON 结构契约作为系统提示词（System Prompt）。
2. 发送至目标大模型接口（默认 `temperature: 0.2`），要求纯 JSON 输出。
3. `extract_json_from_text` 对 AI 返回进行二次清洗（支持提取代码块或内联 JSON 字符串），并校验是否遗漏必填变量字段。

### 4.4 异步任务记录与持久化机制
- 内存持有 `generation_records: Dict[str, Dict]` 缓存，使用 `threading.Lock` 保护并发读写。
- 任务状态流转：`pending` -> `processing` -> `completed` / `failed`。
- 每次任务变更触发 `dump_generation_records()` 将记录按时间倒序同步保存到本地 `generation_records.json`。
- 前端定时轮询 `/generation-records` 获取最新进度，并支持直接下载与一键重试填入提示词。

### 4.5 Excel 批量处理与容错
- 解析参数支持**列字母**（如 `A`、`B`）或**表头字段名**（如 `prompt_text`、`output_name`）。
- 提示词支持宏替换：在批量生成中支持通过 `{{excel_value}}` 或 `{excel_value}` 动态插值。
- 单行处理具备最多 5 次调用重试机制；多次失败后，自动使用空白占位字典回退生成标注 `_fail.docx`，保证批量任务整体不中断。
- 处理完成后统一打包至 `zip` 并提供下载路径。

---

## 5. 开发与运行指南

### 5.1 环境安装
```bash
# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境 (Windows)
.venv\Scripts\activate
# 激活虚拟环境 (Linux/macOS)
source .venv/bin/activate

# 安装核心依赖
pip install -r requirements.txt
```

### 5.2 本地运行
```bash
# 方式 1：直接运行启动脚本（默认端口 8010）
python run_usdoc.py

# 方式 2：使用 uvicorn 开发热重载启动
uvicorn main:app --host 0.0.0.0 --port 8010 --reload
```
服务启动后访问：`http://127.0.0.1:8010/`

### 5.3 Windows 可执行程序打包
```bash
# 安装 PyInstaller
pip install pyinstaller

# 使用预置 spec 打包
pyinstaller --noconfirm usdoc.spec

# 或者使用单行命令打包
pyinstaller --noconfirm --clean --onefile run_usdoc.py --name usdoc --add-data "templates;templates" --add-data "static;static"
```
生成产物位于 `dist/usdoc.exe`。

---

## 6. 接口速查表 (API Specifications)

| 路径 | 方法 | 描述 | 主要请求参数 |
| :--- | :---: | :--- | :--- |
| `/` | `GET` | 页面首页渲染 | - |
| `/upload-template` | `POST` | 上传 `.docx` 模板文件 | `file: UploadFile` |
| `/generate-docx` | `POST` | 提交 Word 异步生成任务 | JSON: `prompt`, `api_url`, `api_key`, `model`, `template_name`, `output_filename` |
| `/generation-records` | `GET` | 查询历史任务记录列表 | Query: `limit` (默认 50) |
| `/generation-records/{id}` | `GET` | 查询单个任务详细状态 | Path: `id` |
| `/batch-tasks` | `POST` | 上传 Excel 异步并发批量生成任务 | Form: `file`, `prompt`, `api_url`, `api_key`, `model`, `template_name`, `sheet_name`, `prompt_column`, `filename_column`, `output_prefix`, `concurrency` |
| `/batch-tasks/{batch_id}` | `GET` | 查询批量任务实时进度、指标、日志及各项明细 | Path: `batch_id` |
| `/batch-tasks/{batch_id}/retry-failed` | `POST` | 一键并发重试全部失败任务项 | Path: `batch_id` |
| `/batch-tasks/{batch_id}/items/{item_id}/retry` | `POST` | 单独重试指定任务行 | Path: `batch_id`, `item_id` |
| `/generate-docx-from-excel` | `POST` | 上传 Excel 批量生成 Word 并打包 ZIP（兼容传统同步接口） | Form: `file`, `api_url`, `api_key`, `model`, `template_name`, `sheet_name`, `prompt_column`, `filename_column`, `prompt`, `output_prefix` |
| `/generate-pdf` | `POST` | 同步生成并转为 PDF 文件 | 同 `/generate-docx` 参数 |
| `/download/{filename}` | `GET` | 文件安全下载（Word/PDF/ZIP） | Path: `filename` |

---

## 7. Agent 维护与编码守则

1. **文件名安全与中文字符支持**：
   - 严禁盲目使用 `re.sub(r"[^a-zA-Z0-9._-]", "_", ...)` 清理用户自定义输出文件名，这会导致中文字符丢失。
   - 必须使用 `sanitize_output_filename` 保留 Unicode 字母和汉字（`re.UNICODE`），仅对目录穿越符号和非法系统路径分隔符做过滤。
2. **模板引擎安全与健壮性**：
   - 当扩展模板功能时，确保对 `TemplateSyntaxError` 做良好捕获，并转化为清晰易懂的 HTTP 400 提示（指示用户占位符书写规则）。
3. **状态与并发锁**：
   - 所有涉及 `generation_records` 字典以及其落盘序列化的操作，必须在 `generation_records_lock` 临界区内进行，防止多协程或后台线程同时操作引发并发数据竞争。
4. **资源清理**：
   - 上传的临时文件（如批量生成中的 `temp_excel_path`）必须在 `finally` 块中确保 `unlink(missing_ok=True)` 且正确 `workbook.close()`，防止句柄泄露与磁盘爆满。
5. **打包兼容性保持**：
   - 修改模板目录或静态资源结构时，必须同步更新 `usdoc.spec` 及 `.github/workflows/build-windows-release.yml` 中的 `--add-data` 参数。
