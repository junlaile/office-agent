# Office Agent

基于 **LangGraph ReAct agent + DeepSeek tool_calls + OfficeCLI** 的交互式 Office 生成 Agent（Word / Excel / PowerPoint）。

参考 [DeepSeek tool_calls 文档](https://api-docs.deepseek.com/zh-cn/guides/tool_calls)，用 LangGraph 原生写法（`bind_tools` + agent loop）：按需求自动识别格式并绑定对应工具集，LLM 自主调用 officecli，直到宣告完成。缺关键信息时会暂停向用户确认。

## 架构

```
用户需求 → 自动识别格式(docx/xlsx/pptx)
   ↓
[agent 节点] ←──────┐
  DeepSeek + bind_tools(该格式工具集)
  自主决定调哪个工具
   ↓
[tools 节点] ──→ 执行(officecli subprocess / interrupt / finish)
  ToolMessage 回传
   ↓ (路由)
  还有 tool_calls → 回 agent
  finish / 无 tool_calls → END
```

- **格式识别**：报表/台账等 → Excel；PPT/幻灯/演示文稿 → PowerPoint；默认 Word（仅「表格」不判为 Excel）
- **LangGraph** 编排 ReAct 循环 + HITL（`interrupt` / `Command(resume=...)`）
- **OfficeCLI** 负责 `.docx` / `.xlsx` / `.pptx` 读写

## 工具集

按会话格式只暴露对应工具（共用 `ask_user` / `finish` / `query_vehicle` / `validate_doc`）。

### Word（.docx）

| 工具 | 作用 |
|---|---|
| `create_doc` | 创建空白 docx（默认 zh-CN） |
| `add_title` / `add_heading` / `add_paragraph` / `add_list_item` | 标题与正文 |
| `add_table` / `add_image` | 表格 / 图片 |
| `add_header` / `add_footer` / `add_page_break` | 页眉页脚 / 分页 |
| `replace_text` / `batch_add` | 查找替换 / 批量写入 |
| `view_outline` / `view_text` | 自查 |

### Excel（.xlsx）

| 工具 | 作用 |
|---|---|
| `create_workbook` | 创建空白工作簿 |
| `add_sheet` | 新增工作表 |
| `write_range` | 从 A1 等起点批量写二维数据（表头可加粗） |
| `write_cell` | 单格值或公式（如 `SUM(B2:B10)`） |
| `view_sheet` | 自查工作簿内容 |

### PowerPoint（.pptx）

| 工具 | 作用 |
|---|---|
| `create_presentation` | 创建空白演示文稿 |
| `add_slide` | 新增一页（标题/正文） |
| `add_bullets` | 页内要点列表 |
| `add_slide_table` / `add_slide_image` | 页内表格 / 图片 |
| `view_outline` / `view_text` | 自查 |

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
# 或直接带需求（自动识别格式）：
python main.py "写一份项目周报，包含本周进展、风险、下周计划"   # → .docx
python main.py "做一份销售月度报表，含区域和金额"                 # → .xlsx
python main.py "做一份项目汇报 PPT，3 页"                        # → .pptx
```

终端会彩色展示 agent 每一步的工具调用（`🔧`）和结果（`↳`）；遇到歧义会暂停问你（`❓`，可输候选序号或自由作答）。生成的文件落在 `output/` 目录。

### 忙时交互

Agent 在推理 / 写文档时（未弹出 `ask_user`），仍可直接打字回车：

| 输入 | 行为 |
|---|---|
| 普通文字 | **软补充**：排队，下一轮推理时注入，Agent 吸收后继续 |
| `!内容` / `/force 内容` / `强制:内容` | **强制打断**：当前节点结束后打断，理解新信息后继续 |
| `继续` / `continue` / `请继续完成` | 暂停后恢复，基于已有文档接着做 |
| `退出` / `quit` | 结束本次运行 |
| `Ctrl+C` | **软暂停**（再按一次或输入退出才真正结束） |

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
│   ├── prompts.py               # 分格式系统提示词（Word/Excel/PPT）
│   ├── format_detect.py         # 需求 → docx/xlsx/pptx
│   ├── state.py                 # 极简 state（messages + doc_path + done）
│   ├── user_input.py            # 忙时 stdin 桥（补充 / 强制打断 / 继续）
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

- 图片、页眉页脚、分页、查找替换、batch_add 已内置；如需 TOC/水印/超链接等，在 `tools.py` 加 `@tool` + `officecli.py` 的 DocTool 加对应方法。
- 接 Web UI：`main.py` 的 interrupt/resume 循环可替换为 WebSocket / HTTP 端点。

## 许可

本项目代码 MIT。OfficeCLI 二进制为 Apache-2.0（iOfficeAI），由脚本独立下载。
