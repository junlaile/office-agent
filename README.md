# Office Agent

基于 **LangGraph ReAct agent + DeepSeek tool_calls + OfficeCLI** 的交互式 Office 文档生成 Agent，支持 **Word / Excel / PowerPoint** 三种格式。

参考 [DeepSeek tool_calls 文档](https://api-docs.deepseek.com/zh-cn/guides/tool_calls)，用 LangGraph 原生写法（`bind_tools` + agent loop）：LLM 自主决定调用哪个 officecli 工具、什么参数，agent 循环自动执行工具并把结果回传，直到 LLM 宣告完成。缺关键信息时会暂停向用户确认。

## 架构

```
用户需求
   ↓
[文档类型推断]（关键词 → .docx/.xlsx/.pptx，拿不准问用户）
   ↓
[agent 节点] ←──────┐
  DeepSeek + bind_tools(按会话类型裁剪的工具子集：通用+格式专属+控制)
  系统提示词按文档类型走对应分支（Word/Excel/PPTX）
  自主决定调哪个工具
  节点边界注入忙时用户补充 / 强制打断
   ↓
[tools 节点] ──→ 执行(officecli subprocess / interrupt / finish)
  _tool() 工厂按扩展名路由到 DocTool/ExcelTool/PptxTool
  连续"末尾追加"类调用合并为一次 officecli batch（失败原子回滚+回退逐个）
  ToolMessage 回传
   ↓ (路由)
  还有 tool_calls → 回 agent
  finish / 无 tool_calls → END
```

- **LangGraph** 编排 ReAct 循环 + 原生 human-in-the-loop（`interrupt` / `Command(resume=...)`）
- **DeepSeek** 通过 OpenAI 兼容 `tools` 接口做工具调用（`bind_tools`，不强制 `tool_choice`）
- **OfficeCLI**（iOfficeAI）负责 docx/xlsx/pptx 实际读写，每个操作暴露为一个 `@tool`
- **UserInputBridge** 单一 stdin 读线程：忙时可补充 / 强制打断，空闲时供 `ask_user` 使用

## 工具集（49 个）

LLM 可自主调用以下工具（无需传文件路径，由会话注入）。文档类型决定哪些工具可用——
每个会话只把"通用 + 对应格式专属 + 控制"的子集 `bind_tools` 给 LLM（`tools_for_kind`），
`query_vehicle` 仅在需求与车辆/交通相关时附加：

### 通用工具（三格式共用，6 个）

| 工具 | 作用 |
|---|---|
| `create_doc` | 创建空白文档（docx/xlsx/pptx） |
| `add_table` | 加表格（Word 末尾 / Excel 当前表 / PowerPoint 最新页） |
| `add_image` | 插入图片（Word 文档末尾 / PowerPoint 最新页） |
| `view_text` | 读文档纯文本（自查） |
| `validate_doc` | 校验 OpenXML 规范 |
| `set_doc_properties` | 设置文档属性（标题/作者/主题/关键词） |

### Word 专属（11 个）

| 工具 | 作用 |
|---|---|
| `add_title` / `add_heading` | 主标题 / 章节标题 |
| `add_paragraph` | 正文段落 |
| `add_list_item` | 列表项（有序/无序） |
| `add_toc` | 插入目录（自动收录标题） |
| `add_page_number` | 页脚页码 |
| `add_header` / `add_footer` | 页眉 / 页脚文字 |
| `add_hyperlink` | 外部超链接 |
| `add_word_chart` | 嵌入式图表（自带数据） |
| `add_section_break` | 分节符（切换横/纵向） |

### Excel 专属（17 个）

| 工具 | 作用 |
|---|---|
| `add_sheet` / `rename_sheet` | 添加 / 重命名工作表 |
| `set_cell` / `set_cells` | 写单个 / 批量写二维数据 |
| `set_formula` | 写公式（不带 `=`） |
| `set_column_width` / `autofit_column` | 列宽 / 自动列宽 |
| `merge_cells` | 合并单元格 |
| `add_excel_chart` | 加图表（柱形/折线/饼图，引用单元格） |
| `add_list_table` | 转真 Excel 表格（带样式+筛选） |
| `set_autofilter` / `sort_sheet` | 自动筛选 / 排序 |
| `highlight_cells` / `add_color_scale` / `add_data_bar` | 条件格式（高亮/色阶/数据条） |
| `add_pivot_table` | 透视表（按字段汇总） |
| `add_dropdown` | 下拉列表（数据验证） |

### PowerPoint 专属（8 个）

| 工具 | 作用 |
|---|---|
| `add_slide` | 添加幻灯片（标题+正文占位符模式） |
| `add_textbox` / `add_slide_table` / `add_slide_image` | 高级：文本框/表格/图片（仅空白页） |
| `set_slide_transition` | 切换效果（fade/morph/push 等） |
| `set_slide_notes` | 演讲者备注 |
| `set_theme_colors` / `set_theme_fonts` | 主题色 / 主题字体 |

### 业务专项 + 控制（3 个）

| 工具 | 作用 |
|---|---|
| `query_vehicle` | 按车牌查询车辆信息（交通类文档用，含照片/事故/违法） |
| `ask_user` | 缺关键信息时向用户提问（触发 interrupt） |
| `finish` | 宣告完成 |

## 快速开始

### 1. 安装依赖

```bash
uv sync
```

### 2. 下载 OfficeCLI 二进制（工程内，跨平台）

```bash
python scripts/fetch_officecli.py
```

自动识别 Windows/Linux/macOS + x64/arm64，下载到 `bin/`，带 SHA256 校验。

下载后建议先验证：

```bash
# Windows
.\bin\officecli.exe --version
# Linux/macOS
./bin/officecli --version
```

若出现 `System.Private.Xml` / `Could not load file or assembly` / `FileNotFoundException`：
多半是 **officecli 缺少匹配的 .NET 运行时**，或下载到的发布包不完整。请：
1. 重新执行 `python scripts/fetch_officecli.py`；
2. Windows 上安装匹配的 [.NET Desktop Runtime](https://dotnet.microsoft.com/download)；
3. 或设置 `OFFICECLI_BIN` 指向已知可用的 OfficeCLI 版本。

> 网络受限？设置代理后重试：
> - PowerShell: `$env:HTTPS_PROXY='http://127.0.0.1:7890'`
> - bash: `export HTTPS_PROXY=http://127.0.0.1:7890`

### 3. 配置 LLM

```bash
cp .env.example .env
# 编辑 .env，填入 LLM_API_KEY
```

默认配置是 DeepSeek；切换其他兼容接口（GLM/OpenAI/通义等）改 `.env` 或 `pyproject.toml` 的 `[tool.office-agent]`。优先级：**环境变量 > .env > pyproject.toml**。

### 4. 运行

```bash
uv run office-agent
# 或：
uv run python -m office_agent
# 兼容 shim：
python main.py

# 直接带需求：
uv run office-agent "写一份项目周报，包含本周进展、风险、下周计划"
uv run office-agent "做一份季度销售数据的 Excel 表格，含图表"
uv run office-agent "做一个 10 页的产品介绍 PPT"
```

启动时会按需求关键词推断文档类型（excel/ppt/word 关键词命中数最高的胜出；无明确线索时交互问一句）。**Word（含法定公文）** 会先生成结构化 Markdown 大纲预览，你可批准、提修改意见循环修订，或取消；**批准后才**创建/合并模板并开始写 `.docx`。Excel / PPT 无此门控。终端会彩色展示 agent 每一步的工具调用（`🔧`）和结果（`↳`）；遇到歧义会暂停问你（`❓`）。生成的文档落在 `output/` 目录。

**Word 大纲预览**（写文件前）：

| 选项 | 含义 |
|---|---|
| `1` / 批准 | 按该大纲生成 Word |
| `2` / 修改 | 输入修改意见后重新生成大纲 |
| `3` / 取消 | 不写文件，结束本次运行 |

**忙时输入**（agent 正在跑时也可键入）：

| 输入 | 含义 |
|---|---|
| 普通文字 | 软补充，下一轮注入给 agent |
| `!内容` / `强制:内容` | 强制打断，节点边界注入 |
| `继续` | 继续当前任务（不重新 create_doc） |
| `退出` | 结束本次运行 |

### 5. Web API（供前端对接）

仅后端 API，覆盖完整对话能力（文档类型确认、Word 大纲批准、公文版头、`ask_user` / `finish` 确认含正文预览、忙时补充/强制打断、文档下载）。

```bash
uv run office-agent-api
# 或: uv run python -m office_agent.api
# 默认 http://0.0.0.0:8000 ；Swagger: /docs
```

| 端点 | 说明 |
|---|---|
| `GET /health` | 健康检查（LLM / officecli） |
| **`POST /v1/chat/completions`** | **OpenAI 兼容对话**（`stream` 可选） |
| `GET /v1/models` | OpenAI 兼容模型列表 |
| `WS /api/v1/ws` | 原生 WebSocket 对话主通道 |
| `GET /api/v1/sessions/{id}` | 会话状态 |
| `GET /api/v1/sessions/{id}/download` | 下载生成的文档 |

#### OpenAI 兼容用法

任意支持 OpenAI API 的客户端，将 Base URL 设为 `http://localhost:8000/v1`，模型名 `office-agent`：

```bash
curl http://localhost:8000/v1/chat/completions \
  -H "Authorization: Bearer any" \
  -H "Content-Type: application/json" \
  -d '{"model":"office-agent","messages":[{"role":"user","content":"写一份项目周报"}]}'
```

多轮续接：助手回复含 `<!--office-agent-session:UUID-->`；也可传请求头 `X-Session-Id`。按提示回复「批准 / 修改… / JSON 表单」即可完成大纲、版头、`ask_user`、finish 确认。

WebSocket 客户端消息：`start` / `choose_kind` / `outline_decision` / `official_header` / `resume` / `supplement` / `force` / `continue` / `quit` / `pause`。

服务端事件：`session` / `need_kind` / `outline` / `need_official_header` / `agent_step` / `tool_result` / `interrupt`（含 `type=confirm_finish` 与 `content_preview`）/ `done`（含 `download_url`）/ `error` / `cancelled`。

## 配置项

| 变量 | 位置 | 说明 |
|---|---|---|
| `LLM_BASE_URL` | pyproject / .env | OpenAI 兼容接口（默认 DeepSeek） |
| `LLM_API_KEY` | **.env**（敏感，勿提交） | API key |
| `LLM_MODEL` | pyproject / .env | 模型名（默认 `deepseek-v4-flash`） |
| `OFFICECLI_BIN` | .env | 二进制路径（留空自动查找） |
| `OUTPUT_DIR` | pyproject / .env | 文档输出目录（默认 `./output`） |
| `SESSION_BACKEND` | .env | 会话存储后端：`memory`（默认）/ `mysql` |
| `MYSQL_HOST`/`MYSQL_PORT`/`MYSQL_USER`/`MYSQL_PASSWORD`/`MYSQL_DATABASE` | .env | `SESSION_BACKEND=mysql` 时生效 |

## 目录结构

```
office-agent/
├── main.py                      # 兼容 shim → office_agent.cli.main
├── 新增office模版的说明.md        # 怎么加一个新公文模板
├── scripts/
│   ├── fetch_officecli.py       # 跨平台下载 officecli
│   └── gen_official_templates.py # 生成 15 文种公文模板（GB/T 9704）
├── src/office_agent/
│   ├── config.py                # 配置加载（env > .env > pyproject）
│   ├── officecli.py             # 向后兼容门面（re-export office.*）
│   ├── agent/                   # LangGraph 核心
│   │   ├── graph.py / state.py / prompts.py / llm.py
│   │   └── outline.py           # Word 预览大纲生成（批准前）
│   ├── cli/                     # 终端入口与交互
│   │   ├── main.py              # run() 编排 + UserInputBridge 生命周期
│   │   ├── ui.py                # 终端 UI / interrupt 表单
│   │   └── user_input.py        # 忙时输入桥（补充/强制/继续/退出）
│   ├── session/                 # CLI/API 共用会话编排
│   │   ├── prep.py              # 路径/版头/模板 merge（无 UI）
│   │   └── runner.py            # AgentSession 状态机
│   ├── api/                     # FastAPI WebSocket + 下载
│   │   ├── app.py               # 路由与 WS 协议
│   │   └── manager.py           # 进程内会话表
│   ├── office/                  # OfficeCLI 实现层
│   │   ├── runner.py            # subprocess 执行器 + merge_template
│   │   ├── doc.py / excel.py / pptx.py   # DocTool / ExcelTool / PptxTool
│   ├── domain/                  # 业务域
│   │   ├── templates.py         # 15 文种元数据 + 文种识别
│   │   ├── format.py            # Office 格式推断（docx/xlsx/pptx）
│   │   └── vehicle_data.py      # 车辆查询 mock
│   └── tools/                   # @tool 工具集包
│       ├── __init__.py          # 会话基础设施 + ALL_TOOLS 聚合
│       ├── common.py / doc.py / excel.py / pptx.py
├── template/word/               # 15 文种公文模板（GB/T 9704）
├── tests/                       # 测试套件（pytest）
│   ├── conftest.py              # FakeRunner + session fixture
│   ├── test_*.py                # 单元测试（默认跑）
│   ├── integration/             # 集成测试（@pytest.mark.integration，默认 skip）
│   └── llm/                     # LLM 端到端（@pytest.mark.llm，默认 skip）
├── bin/                         # officecli 二进制（gitignore）
└── output/                      # 生成的文档（gitignore）
```

## 设计要点

- **DeepSeek tool_calls**：`llm.bind_tools(tools)`；不强制 `tool_choice`，不用 `response_format`（thinking 模式冲突）。
- **Word 大纲预览门控**：写 `.docx` 前用无工具 LLM 生成 Markdown 大纲；用户批准后才 `_prepare_official_doc` / `create_doc`。
- **手写 tools 节点**：比内置 `ToolNode` 灵活，能处理 `ask_user` 的 `interrupt` 挂起和 `finish` 的短路完成。
- **会话级 doc_path 注入**：启动时按需求推断文档类型和输出路径，所有工具共享，LLM 不需传路径参数。
- **扩展名路由**：`tools._tool()` 工厂按 `.docx/.xlsx/.pptx` 后缀返回对应 Tool 类；通用工具跨格式，专属工具在其他格式下给出明确引导。
- **公文模式**：识别 15 法定文种 → 从 GB/T 9704 模板创建 → LLM 用 update_paragraph/replace_text/remove_paragraph 编辑正文（保字体）。
- **忙时输入桥**：单一 stdin 读线程；忙时分类投递，空闲时 `blocking_readline` 供 ask_user，避免抢 stdin。
- **finish 显式结束**：比"无 tool_calls 即结束"更可靠。
- **中文 UTF-8**：subprocess 强制 `encoding='utf-8'`，JSON 通过 `--commands` argv 传递。
- **显式格式 props**：默认 docx 缺 Heading/Title 样式，用 `size`/`bold` 等显式 props 确保格式生效。
- **门面模式保兼容**：`officecli.py` 保留为 re-export 门面，`from office_agent.officecli import X` 零改动。

## 开发

### 测试

```bash
# 单元测试（默认，毫秒级，全 mock 不碰外部）
uv run pytest

# 集成测试（真调 officecli.exe，需 Windows + Office）
uv run pytest -m integration

# LLM 端到端（真调 API，耗 token）
uv run pytest -m llm

# 全跑（含集成 + LLM）
uv run pytest -m ""

# 覆盖率报告
uv run pytest --cov-report=html  # 生成 htmlcov/
```

测试分层（详见 `tests/`）：
- **单元测试**（默认跑）：纯函数 + FakeRunner mock，覆盖 domain/cli/office/tools/agent。覆盖率门槛 65%。
- **集成测试**（`@pytest.mark.integration` 默认 skip）：真调 officecli.exe，验证 create/add/view/merge/validate 端到端。
- **LLM 端到端**（`@pytest.mark.llm` 默认 skip）：真跑 agent 生成公文，验证完整链路。

### 质量工具

```bash
uv run ruff format .        # 格式化
uv run ruff check .         # lint
uv run mypy src/            # 类型检查
```

### 扩展

- **新增公文模板**（加一个文种）：见[《新增office模版的说明.md》](新增office模版的说明.md)。只需在 `domain/templates.py` 的注册表登记 + 在 `scripts/gen_official_templates.py` 写正文范例，其余（文件名、提示词文种清单/结语、工具描述、测试）自动派生。
- 如需更多 officecli 能力（条件格式、数据透视表、幻灯片切换动画、SmartArt 等），在 `office/doc.py` / `excel.py` / `pptx.py` 对应类加方法 + `tools/doc.py` / `excel.py` / `pptx.py` 加 `@tool`。命令面参考 `officecli help <format> <element>`。
- 接 Web UI：使用 `session.AgentSession` + `api` WebSocket 协议；前端只需实现事件驱动的对话页。

## 许可

本项目代码 MIT。OfficeCLI 二进制为 Apache-2.0（iOfficeAI），由脚本独立下载。
