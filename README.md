# Office Agent

基于 **LangGraph ReAct agent + DeepSeek tool_calls + OfficeCLI** 的交互式 Word 文档生成 Agent。

参考 [DeepSeek tool_calls 文档](https://api-docs.deepseek.com/zh-cn/guides/tool_calls)，用 LangGraph 原生写法（`bind_tools` + agent loop）：LLM 自主决定调用哪个 officecli 工具、什么参数，agent 循环自动执行工具并把结果回传，直到 LLM 宣告完成。缺关键信息时会暂停向用户确认。

## 架构

```
用户需求
   ↓
[agent 节点] ←──────┐
  DeepSeek + bind_tools([
    create_doc, add_title, add_heading, add_paragraph,
    add_list_item, add_table, view_text, validate_doc,
    ask_user, finish
  ])
  自主决定调哪个工具
   ↓
[tools 节点] ──→ 执行(officecli subprocess / interrupt / finish)
  ToolMessage 回传
   ↓ (路由)
  还有 tool_calls → 回 agent
  finish / 无 tool_calls → END
```

- **LangGraph** 编排 ReAct 循环 + 原生 human-in-the-loop（`interrupt` / `Command(resume=...)`）
- **DeepSeek** 通过 OpenAI 兼容 `tools` 接口做工具调用（`bind_tools`，不强制 `tool_choice`）
- **OfficeCLI**（iOfficeAI）负责 docx 实际读写，每个操作暴露为一个 `@tool`

## 工具集

LLM 可自主调用以下工具（无需传文件路径，由会话注入）：

| 工具 | 作用 |
|---|---|
| `create_doc` | 创建空白 docx |
| `add_title` / `add_heading` | 主标题 / 章节标题 |
| `add_paragraph` | 正文段落 |
| `add_list_item` | 列表项（有序/无序） |
| `add_table` | 表格（二维数组） |
| `view_text` / `validate_doc` | 自查内容 / 校验文档 |
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
python main.py
# 或直接带需求：
python main.py "写一份项目周报，包含本周进展、风险、下周计划"
```

终端会彩色展示 agent 每一步的工具调用（`🔧`）和结果（`↳`）；遇到歧义会暂停问你（`❓`，可输候选序号或自由作答）。生成的文档落在 `output/` 目录。

## 配置项

| 变量 | 位置 | 说明 |
|---|---|---|
| `LLM_BASE_URL` | pyproject / .env | OpenAI 兼容接口（默认 DeepSeek） |
| `LLM_API_KEY` | **.env**（敏感，勿提交） | API key |
| `LLM_MODEL` | pyproject / .env | 模型名（默认 `deepseek-v4-flash`） |
| `OFFICECLI_BIN` | .env | 二进制路径（留空自动查找） |
| `OUTPUT_DIR` | pyproject / .env | 文档输出目录（默认 `./output`） |

## 目录结构

```
office-agent/
├── main.py                      # 终端入口（ReAct 交互循环）
├── scripts/fetch_officecli.py   # 跨平台下载 officecli
├── src/office_agent/
│   ├── config.py                # 配置加载（env > .env > pyproject）
│   ├── llm.py                   # ChatOpenAI 工厂
│   ├── officecli.py             # OfficeCLI subprocess 封装 + DocTool
│   ├── tools.py                 # @tool 工具集 + 会话注入
│   ├── prompts.py               # ReAct agent 系统提示词
│   ├── state.py                 # 极简 state（messages + doc_path + done）
│   └── graph.py                 # ReAct agent 图装配
├── bin/                         # officecli 二进制（gitignore）
└── output/                      # 生成的 docx（gitignore）
```

## 设计要点

- **DeepSeek tool_calls**：`llm.bind_tools(tools)`（实测可行）；注意 DeepSeek 不支持 `response_format`（json schema）和强制 `tool_choice`（thinking 模式冲突），所以用普通 `bind_tools` 而非 `with_structured_output`。
- **手写 tools 节点**：比内置 `ToolNode` 灵活，能处理 `ask_user` 的 `interrupt` 挂起和 `finish` 的短路完成。
- **会话级 doc_path 注入**：main.py 启动时确定输出路径，所有工具共享，LLM 不需传路径参数。
- **finish 显式结束**：比"无 tool_calls 即结束"更可靠，LLM 主动宣告完成。
- **中文 UTF-8**：subprocess 强制 `encoding='utf-8'`，JSON 通过 `--commands` argv 传递（stdin 传中文会乱码）。
- **显式格式 props**：默认 docx 缺 Heading/Title 样式，用 `size`/`bold`/`listStyle` 等显式 props 确保格式生效。

## 扩展

- 支持表格/列表已内置；如需图片、页眉页脚，在 `tools.py` 加 `@tool` + `officecli.py` 的 DocTool 加对应方法。
- 接 Web UI：`main.py` 的 interrupt/resume 循环可替换为 WebSocket / HTTP 端点。

## 许可

本项目代码 MIT。OfficeCLI 二进制为 Apache-2.0（iOfficeAI），由脚本独立下载。
