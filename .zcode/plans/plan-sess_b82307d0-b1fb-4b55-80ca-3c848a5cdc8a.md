## 问题根因

"图片不可下载或引用不存在的情况下不嵌入图片"。实测 officecli 对坏源**已经会报错**（404/文件不存在/空源都返回非零），所以**不会**嵌入坏图。但存在两个问题：

1. **悬空空段落**（主问题）：`DocTool.add_image`（`doc_tool.py:279-300`）先 `add paragraph` 建载段，再 `add picture`。图片失败时载段**已留在文档里**。实测：3 次失败尝试 → 文档里 3 个空段落、0 张图片。
2. **错误提示无指导性**：工具层捕获 `OfficeCLIError` 返回 `插入图片失败: officecli 返回非零状态...`，又长又没告诉 LLM "跳过、别重试"，LLM 可能反复重试堆积更多空段。

## 修复方案：预校验图片来源，无效则在碰文档前跳过

核心：在 `common.py` 新增 `_validate_image_source(src)`，工具层在分发到 DocTool/PptxTool **之前**先校验。无效就返回明确跳过提示、**完全不碰文档**（不建载段、不调 officecli）。

### 改动 1：`src/office_agent/tools/common.py` — 新增 `_validate_image_source`
- 用标准库 `urllib.request`（无新依赖）。
- 返回 `None` = 有效；返回 `str` = 失败原因（给 LLM 看）。
- 三类源的处理：
  - **空串** → 失败"图片来源为空"。
  - **data URI**（`data:` 开头）→ 直接放行（内联数据，总是可用）。
  - **本地路径**（非 http）→ `os.path.exists` 检查，不存在则失败"文件不存在: {src}"。可靠、无歧义，坚决拦。
  - **HTTP/HTTPS URL** → 发 HEAD 请求（`urllib.request.Request(method="HEAD")`，超时 8s）。
    - 明确 **404/410**（Not Found）→ 拦，失败"图片不存在（HTTP {code}）"。
    - **连接错误/DNS 失败/超时** → 拦，失败"图片不可访问（{简短原因}）"。
    - **405（方法不允许）/其它非确定性响应** → **不拦**（放行交给 officecli 实际 GET），避免误伤（有些服务器不支持 HEAD 但 GET 能取）。
    - **HEAD 成功（2xx/3xx）** → 放行。
- 设计为**宁可放行不可误拦**：HEAD 不确定时让 officecli 兜底（officecli 失败时工具层原有 try/except 仍会捕获，只是会留空段——但这是边缘情况，主流坏源 404/DNS 失败/本地不存在已被预校验拦住）。

### 改动 2：`src/office_agent/tools/common.py` — `add_image` 工具开头加预校验
```python
reason = _validate_image_source(url_or_path)
if reason:
    return (
        f"⚠️ 跳过插入图片（{reason}）。"
        f"不要重试这张图，继续生成文档其他内容。"
    )
```
放在 `try` 之前，确保无效时不进 try、不碰文档。

### 改动 3：`src/office_agent/tools/pptx.py` — `add_slide_image` 同样加预校验
- 复用 `_validate_image_source`（从 common 导入），同样的跳过提示。PPTX 路径虽无悬空段问题（PptxTool.add_image 是单次 add picture，无载段），但预校验仍能避免无谓的 officecli 调用、给 LLM 更清晰的反馈。

### 改动 4：`src/office_agent/doc_tool.py` — `add_image` 防御性兜底（事后清理）
- 即使预校验放行，officecli 仍可能因边缘情况失败（如 HEAD 放行但实际 GET 404）。给 DocTool.add_image 加兜底：用 try 包住"建载段+插图"，失败时 `remove` 掉刚建的载段再上抛。
- 双保险：预校验拦主流坏源（不碰文档）+ 事后清理兜底边缘情况（不留空段）。

### 改动 5：测试
- `tests/test_image_validation.py`（新增，单测，不碰 officecli）：
  - 空串 → 失败原因含"为空"。
  - data URI → 返回 None（有效）。
  - 本地不存在路径 → 失败原因含"不存在"。
  - 本地存在路径 → None（用 tmp_path 建临时文件）。
  - HTTP 404（mock urllib）→ 失败原因含"不存在"/"404"。
  - HTTP 405（mock）→ None（放行）。
  - HTTP 连接异常（mock）→ 失败原因含"不可访问"。
- `tests/test_tools_forwarding.py`（扩展）：`add_image` 传坏源（不存在的本地路径）→ 返回值含"跳过"、且**不调 runner**（用 fake_runner 断言 calls 为空，验证不碰文档）。

### 为什么这样最稳
- **预校验**：主流坏源（本地不存在、404、DNS 失败）在碰文档前拦掉，零副作用、零空段、省 officecli 调用。
- **事后清理兜底**：预校验放行但 officecli 仍失败的边缘情况（HEAD 不准），删掉载段不留痕迹。
- **清晰反馈**：跳过提示明确告诉 LLM"别重试、继续其他内容"，避免堆积。
- **无新依赖**：纯标准库 urllib。
- **两个图片工具都覆盖**：add_image（docx+pptx）+ add_slide_image（pptx）。

### 验证
- `pytest tests/test_image_validation.py -q`（新增单测）。
- `pytest tests/test_tools_forwarding.py -q`（扩展，不调 runner）。
- `pytest -q`（全量无回归）。
- integration（可选）：真 officecli + 坏 URL，确认无空段残留（标 @pytest.mark.integration）。