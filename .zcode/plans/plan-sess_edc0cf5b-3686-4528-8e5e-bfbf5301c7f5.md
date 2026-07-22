# 修复 PPT 文字重叠(占位符模式)

## 根因(已证实)

1. **画布尺寸认知错误**:提示词说"幻灯片约 24cm × 13cm",实际是 **33.87cm × 19.05cm**(标准 16:9 宽屏,cx=12192000 cy=6858000 EMU)。LLM 按错误尺寸规划坐标。
2. **`add_slide(title, body_text)` 会建一个大正文占位符**:位置 y≈5.08cm 到 y≈17.16cm(高 12.08cm)、x=0.92cm、宽 29.21cm,几乎占满下半幅。LLM 若再在同页加 `add_textbox`(如 y=4cm),必然与之重叠——这就是观察到的重叠。
3. **`layout="Blank"` 不会产生空白页**(实测 + 官方文档确认):占位符只由 `title=`/`text=` 控制,layout 只是元数据。提示词之前暗示 `layout=Blank` 能给空白页,是错的。

## 修复方案:切换为「占位符模式」

你已选定占位符模式:每页用 `add_slide(title, body_text)` 写全部内容,body_text 承载多行文本(换行 `\n` 拆成多个 paragraph),不再在同页叠加 `add_textbox`。

### 改动 1 — `tools.py` 的 `add_slide` 工具 docstring

重写参数说明,明确占位符模式:
- `title`:标题文字(非空 → 建标题占位符,位于页顶)。
- `body_text`:正文(非空 → 建正文占位符,占页中下部)。**这是写正文内容的唯一入口**,支持换行 `\n`、要点符号 `·`/`•`/`-`、空行分段。
- `layout`:降级为"可选元数据,通常留空即可"。删除原来暗示 Blank 能给空白页的误导文案。
- 新增"重要"提示:**每页内容只通过 title + body_text 写入。不要在同一页再调 add_textbox/add_slide_table/add_slide_image——会与正文占位符重叠。** 一页放不下的内容拆成多页。

### 改动 2 — `tools.py` 的三个自由排版工具 docstring 加警告

`add_textbox` / `add_slide_table` / `add_slide_image`:保留(不删,以备高级用途),但在 docstring 顶部加醒目警告:
> ⚠️ 高级工具。常规 PPT 请用 `add_slide(title, body_text)` 即可,【不要】在已有占位符的页上再调本工具(会重叠)。本工具仅用于 `add_slide()` 不带 title/body_text 的纯空白页上的精细排版。

### 改动 3 — `prompts.py` 的 `_PPTX_BRANCH` 重写

- 修正画布尺寸:**33.87cm × 19.05cm**(16:9)。
- 工作流改为纯占位符模式:
  ```
  1. 规划:分几页?每页标题 + 要点内容。
  2. create_doc
  3. 逐页 add_slide(title, body_text):
     - 封面: add_slide(title='主题', body_text='副标题/日期/汇报人')
     - 内容页: add_slide(title='本页主题', body_text='要点1\n要点2\n...\n\n小节:\n· 细节A\n· 细节B')
  4. view_text 自查 → finish
  ```
- 明确禁令:**不要混用 add_textbox**;一页放不下就拆页。
- body_text 排版规范:每行一个要点,用 `·`/`-`/`•` 前缀;不同小节之间空行;不要塞太多(每页 5-8 行为宜)。
- 删除"Blank 布局 + add_textbox 自由排版"的引导。

## 不改动

- `officecli.py` 的 `PptxTool` 类:方法本身没错(坐标是 officecli 内置的合理值),重叠源于上层 LLM 误用 + 提示词尺寸错误,不需要改底层。
- `add_textbox` 的默认 width(22cm):占位符模式下不再被 LLM 常用,改不改影响不大;保留以减少变动。但可以顺手把默认 x/width 调整得更贴合 33.87cm 画布(如 width 默认 28cm 居中),作为兜底——可选。

## 验证

1. 跑一次 `python main.py "做一个5页的产品介绍PPT..."`,确认 LLM 只用 `add_slide(title, body_text)`,不再叠加 `add_textbox`,生成的 PPT 无重叠。
2. 用 `get /slide[N] --depth 2` 检查每页只有 2 个 shape(标题+正文),无额外文本框叠加。
3. 确认 body_text 的多行内容正确拆成多个 paragraph、中文无乱码。