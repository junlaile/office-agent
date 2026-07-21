## 目标

对 office-agent 做"中等粒度"拆分重构 + 完整测试覆盖（单测 + 集成 + LLM 三层）+ 配置 ruff/mypy/pytest-cov 质量工具。前提：保持所有现有功能行为不变（公文模式、普通 Word/Excel/PPT 模式都不能回归）。

## 决策（已与用户确认）

- 拆分粒度：中等（officecli → 4 文件；tools → 包；main 纯函数下沉）
- 测试范围：单测（默认跑）+ 集成（officecli.exe，`@pytest.mark.integration` 默认 skip）+ LLM（`@pytest.mark.llm` 默认 skip）
- 质量工具：ruff（lint+format）+ mypy（类型）+ pytest-cov（覆盖率，目标 ≥80%）

## 实施步骤（分 4 阶段）

### 阶段 1：测试脚手架 + 质量工具配置

**1.1 安装开发依赖**
```bash
uv add --dev pytest pytest-cov pytest-asyncio ruff mypy
```

**1.2 pyproject.toml 增补配置**
```toml
[dependency-groups]
dev = ["pytest>=8", "pytest-cov", "pytest-asyncio", "ruff", "mypy"]

[tool.pytest.ini_options]
testpaths = ["tests"]
python_files = ["test_*.py"]
addopts = "-ra --strict-markers --cov=office_agent --cov-report=term-missing --cov-fail-under=80"
markers = [
    "integration: 真调 officecli.exe 的慢测试（默认 -m 'not integration' 跳过）",
    "llm: 真调 LLM API 的测试（默认 skip）",
]

[tool.ruff]
line-length = 100
target-version = "py313"
[tool.ruff.lint] select = ["E","F","W","I","UP","B","SIM"]

[tool.mypy]
python_version = "3.13"
strict = false  # 渐进式，先不开 strict
ignore_missing_imports = true  # langchain/langgraph 无类型存根

[tool.coverage.run]
source = ["office_agent", "main"]
omit = ["*/tests/*", "*/scripts/*"]
```

**1.3 新建 tests/ 目录结构**
```
tests/
  conftest.py              # FakeRunner、session fixture、tmp output fixture、markers 注册
  test_templates.py        # 文种识别 + merge 数据
  test_prompts.py          # build_system_prompt 分支
  test_vehicle_data.py     # mock_query 确定性 + 脱敏
  test_officecli_argv.py   # 注入 FakeRunner 断言生成的 argv
  test_officecli_pure.py   # _build_table_ops / _parse_tbl_index / _col_to_letter 纯函数
  test_tools_session.py    # session 路由 + start_from_template 编排（mock merge）
  test_graph_routing.py    # _route_after_* / _tools_node 分支（不碰 LLM）
  test_main_helpers.py     # _infer_doc_kind / _format_tool_call / _indent
  test_official_doc.py     # 公文模式编排集成（mock officecli，不碰 exe）
  integration/
    __init__.py
    test_officecli_smoke.py    # @pytest.mark.integration 真调 exe
    test_official_merge.py     # @pytest.mark.integration 真实模板 merge
  llm/
    __init__.py
    test_e2e_official.py       # @pytest.mark.llm 真跑 agent 生成公文
```

### 阶段 2：代码拆分（中等粒度）

**2.1 拆 officecli.py（1581 行 → 4 文件）**

把 `src/office_agent/officecli.py` 拆成：
- `src/office_agent/officecli.py`（保留，缩成 ~50 行）：**向后兼容门面**，re-export `OfficeCLIError`/`resolve_bin`/`merge_template`/`DocTool`/`ExcelTool`/`PptxTool`/`raw`/`get_runner`/`reset_runner`。这样所有现有 `from .officecli import ...` 不用改。
- `src/office_agent/cli_runner.py`（~200 行）：`OfficeCLIError`、`_Runner`、`get_runner`/`reset_runner`、`resolve_bin`、`raw`、`merge_template`、白名单。基础设施，被三个 *Tool 依赖。
- `src/office_agent/doc_tool.py`（~500 行）：`DocTool` 类（30 方法）。
- `src/office_agent/excel_tool.py`（~500 行）：`ExcelTool` 类（27 方法）+ `_col_to_letter`/`_ref_at` 辅助。
- `src/office_agent/pptx_tool.py`（~400 行）：`PptxTool` 类（21 方法）。

**关键**：`officecli.py` 保留为门面（re-export），保证 `tools.py`/`main.py`/`graph.py` 的 `from .officecli import` 零改动。这是低风险拆分。

**2.2 拆 tools.py（1384 行 → 包）**

把 `src/office_agent/tools.py` 改造为包：
- `src/office_agent/tools/__init__.py`：会话基础设施（`_session_doc_path`/`set_session_doc`/`session_doc_path`/`session_doc_kind`/`_tool`/`_wrong_kind_msg`）+ 聚合 `ALL_TOOLS`/`TOOL_BY_NAME`（从子模块 import）。
- `src/office_agent/tools/common.py`：通用工具（`create_doc`/`add_table`/`view_text`/`validate_doc`/`add_image`/`set_doc_properties`/`start_from_template`/`ask_user`/`finish`/`query_vehicle`）+ `AskField` model。
- `src/office_agent/tools/doc.py`：Word 工具（`add_title`/`add_heading`/`add_paragraph`/.../`update_paragraph`/`replace_text`/`remove_paragraph` 等 14 个）。
- `src/office_agent/tools/excel.py`：Excel 工具（17 个）。
- `src/office_agent/tools/pptx.py`：PPT 工具（8 个）。

**关键约束**：会话基础设施（`_session_doc_path` 等）必须放在 `__init__.py`，子模块 `from . import _tool, session_doc_kind, _wrong_kind_msg` 引用，避免循环依赖。

**2.3 拆 main.py（546 行）**

把 main.py 的纯函数和 UI 辅助下沉到 `src/office_agent/cli_ui.py`：
- `cli_ui.py`：`_format_tool_call`（86 行）、`_print_agent_step`、`_print_tool_results`、`_indent`、`_banner`、`_ask_doc_kind`、`_collect_form`、`_collect_single_question`、`_handle_interrupt`、`_infer_doc_kind`、`_derive_doc_path`、`_prepare_official_doc`、ANSI 颜色常量。
- `main.py`（缩到 ~100 行）：只留 `run()` 主流程编排 + `__main__` 入口，从 `office_agent.cli_ui` import 辅助。

**关键**：`_infer_doc_kind`/`_format_tool_call` 下沉后成为可 import 的纯函数，测试不再受 main.py 的 sys.path 副作用影响。

**2.4 拆 scripts/gen_official_templates.py 的 build_doc（57 行）**

把 `build_doc` 里的"版头/标题/正文/落款/版记"组装抽成独立小函数（`_assemble_header`/`_assemble_body`/`_assemble_record`），build_doc 只做编排。可选，优先级低。

### 阶段 3：编写测试（按可测试性顺序）

**3.1 conftest.py 核心 fixture**
- `fake_runner`：一个 FakeRunner 类，`run(args, json_output=False)` 把 args 存入 `self.calls` 列表，返回可配置的 mock stdout。提供 `reset_runner` 钩子注入。
- `doc_session(tmp_path)`：`set_session_doc(str(tmp_path/"test.docx"))` + yield + teardown 清空全局。
- `clean_session`：teardown 清空 `_session_doc_path`，防测试间污染。
- 注册 `integration`/`llm` marker 自动 skip。

**3.2 纯函数单测（第一批，覆盖核心逻辑）**

`test_templates.py`（~30 用例）：
- `detect_doc_type`：15 文种各 1 个正例 + 干扰负例（Excel/PPT/周报/调研报告）+ 上行文优先级 + 批复特殊优先（"批复下级的请示"→批复）+ 多文种同时命中的优先级排序。
- `default_merge_data`：overrides > defaults > 全局默认；上行文有 signer、非上行文 signer=""；会议类 issuer/date_cn 为空；命令 doc_no="第 X 号"；**断言返回值无 `{{` 残留**。
- `template_path`：合法文种返回正确路径；未知文种抛 ValueError。
- `is_upward`/`is_meeting`：请示/报告/议案 True，其他 False；决议/纪要 meeting True。

`test_prompts.py`（~10 用例）：
- `build_system_prompt(doc_path)` 按 .docx/.xlsx/.pptx 扩展名选对分支（断言含 WORD/EXCEL/POWERPOINT 关键词）。
- `doc_type` 非空时走公文分支：含 `公文模式`、`update_paragraph`、`replace_text`、结语规范；上行文含"上行文特别提示"。
- 普通模式不含公文分支标记。

`test_vehicle_data.py`（~8 用例）：
- `query(同车牌)` 两次调用结果 deep equal（确定性）。
- 空车牌 / 不存在车牌返回 not_found。
- `_mask_id_card`("110101199001011234") → "110101********1234"。
- `_mask_phone`("13800138000") → "138****8000"。

`test_main_helpers.py`（→ 改名 `test_cli_ui.py`，~15 用例）：
- `_infer_doc_kind`：Excel/PPT/Word 关键词命中；平局优先级 xlsx>pptx>docx；无命中 (None,0)。
- `_format_tool_call`：每个工具分支（add_title/add_heading/add_table/add_slide/set_cells/ask_user/finish/update_paragraph/replace_text/remove_paragraph/start_from_template）格式正确。
- `_indent`：多行缩进。

`test_officecli_pure.py`（~10 用例）：
- `DocTool._build_table_ops`：has_header=True 表头加 bold；rows/cols 正确；路径 `/body/tbl[N]/tr[R]/tc[C]`。
- `DocTool._parse_tbl_index`：解析 "Added table at /body/tbl[3]" → 3；无匹配 → 0。
- `ExcelTool._col_to_letter`：1→A、27→AA、702→ZZ。
- `ExcelTool._ref_at`：(2,3)→"C2"。

**3.3 mock runner 的工具层测试（第二批）**

`test_officecli_argv.py`（~25 用例）：注入 FakeRunner，断言生成的 argv 数组正确：
- `DocTool.add_title("标题")` → args 含 `["add", path, "/body", "--type", "paragraph", "--prop", "text=标题", "--prop", "size=26", "--prop", "bold=true", "--prop", "align=center"]`。
- `DocTool.add_heading/add_paragraph/add_list_item/add_table` 各验证 argv。
- `DocTool.set_paragraph_text/find_replace/remove` 新方法 argv 正确（含 `--find`/`--replace`）。
- `ExcelTool.set_cell/set_cells/set_formula` argv。
- `PptxTool.add_slide/add_textbox` argv。
- `merge_template` argv 含 `--data <json>`。

`test_tools_session.py`（~15 用例）：
- `session_doc_kind()`：.docx/.xlsx/.pptx 路由；未初始化抛错。
- `start_from_template`：非 docx 会话返回 wrong_kind；docx 会话下 mock `merge_template` 断言被调用 + merge_data 含 overrides；title/addressee 提供时返回提示信息。
- `update_paragraph`/`replace_text`/`remove_paragraph`：非 docx 返回 wrong_kind；docx 下 mock DocTool 方法断言转发。

**3.4 graph 路由测试（第三批）**

`test_graph_routing.py`（~10 用例）：构造假 AIMessage state，不碰 LLM：
- `_route_after_agent`：有 tool_calls → "tools"；无 tool_calls → END；空 messages → END。
- `_route_after_tools`：done=True → END；done=False → "agent"。
- `_tools_node`：finish 短路（state.done=True + summary）；ask_user 与其他工具同批时两者都被取消（返回引导 ToolMessage）；未知工具返回错误 ToolMessage；正常工具调用转发 + ToolMessage。

**3.5 编排集成测试（mock officecli，不碰 exe）**

`test_official_doc.py`（~8 用例）：mock `merge_template`，验证公文模式编排：
- `_prepare_official_doc("通知", path)` 调 merge_template + 打印 + 返回"通知"。
- 模板缺失时返回 None。
- merge 失败时返回 None。

**3.6 集成测试（默认 skip）**

`integration/test_officecli_smoke.py`（`@pytest.mark.integration`）：真调 exe，create → add_paragraph → view_text → validate，tmp_path 隔离。
`integration/test_official_merge.py`：真实 `template/word/08-通知.docx` 跑 merge_template，断言 `{{org}}` 被替换、无残留。

**3.7 LLM 端到端（默认 skip）**

`llm/test_e2e_official.py`（`@pytest.mark.llm`）：真跑 `python main.py "写一份通知..."`，断言生成文件、validate 通过、无占位残留。复用之前验证过的真实 API 流程。

### 阶段 4：验证 + 收尾

**4.1 全量验证**
```bash
uv run pytest                    # 单测全过 + 覆盖率 ≥80%
uv run pytest -m integration     # 集成测试（本机跑）
uv run ruff check .              # lint 无错
uv run ruff format --check .     # 格式无错
uv run mypy src/                 # 类型检查无错
uv run python main.py "写一份关于做好防汛工作的通知"  # 真实公文生成不回归
```

**4.2 更新文档**
- README.md：补"开发"段落（测试命令、覆盖率、质量工具）
- template/word/README.md：模块表更新（officecli.py → 拆分后的文件）

## 关键设计原则

1. **门面模式保兼容**：`officecli.py` 缩成 re-export 门面，`from .officecli import X` 零改动。低风险。
2. **会话基础设施集中**：tools 包的 `_session_doc_path` 等放 `__init__.py`，子模块反向引用，避免循环。
3. **测试不碰外部**：单测全 mock（FakeRunner / monkeypatch），毫秒级、零外部依赖。集成/LLM 测试默认 skip，本机按需跑。
4. **覆盖率门槛**：`--cov-fail-under=80` 写进 addopts，CI 强制。
5. **渐进式 mypy**：先 `strict=false`，只抓明显类型错误，不阻断。
6. **真实公文回归**：最后用真 LLM 跑一遍通知生成，确保拆分没破坏公文模式。

## 涉及文件

| 类型 | 文件 |
|---|---|
| **新建（拆分目标）** | `cli_runner.py`、`doc_tool.py`、`excel_tool.py`、`pptx_tool.py`、`cli_ui.py`、`tools/__init__.py`、`tools/common.py`、`tools/doc.py`、`tools/excel.py`、`tools/pptx.py` |
| **改造（缩成门面/编排）** | `officecli.py`（门面）、`tools.py`→删除改为包、`main.py`（缩到 ~100 行）、`gen_official_templates.py`（build_doc 拆小） |
| **新建（测试）** | `tests/conftest.py` + 9 个单测文件 + `integration/` 2 个 + `llm/` 1 个 |
| **配置** | `pyproject.toml`（pytest/ruff/mypy/coverage 配置）、`.coveragerc`（可选） |
| **文档** | `README.md`、`template/word/README.md` |

## 不做的事

- 不改任何业务逻辑/公文规范/模板内容（纯结构重构 + 测试）
- 不动 LangGraph 图结构、prompt 文案
- 不动 .env / 真实模板文件
- 不引入新的运行时依赖（只加 dev 依赖）
- mypy 暂不开 strict（避免大量类型修复阻断主任务）

## 风险与应对

- **风险**：tools.py 改包后 import 路径变（`from office_agent.tools import X` 仍可用，因为 `__init__.py` 聚合）。**应对**：__init__.py 显式 re-export 所有公开符号。
- **风险**：拆分后循环 import（tools 子模块 ↔ tools/__init__）。**应对**：会话基础设施只在 __init__，子模块用 `from office_agent.tools import _tool`（延迟或在函数内 import）。
- **风险**：main.py 的 sys.path 副作用影响测试。**应对**：纯函数下沉到 cli_ui.py 后，测试 import cli_ui 不触发 main.py 副作用。
- **风险**：FakeRunner 不能完全模拟 officecli.exe 行为（如 view_text 的输出格式）。**应对**：FakeRunner 允许按命令类型返回预设 stdout，关键测试断言 argv 而非真实执行结果；真实执行留给 integration 测试。
