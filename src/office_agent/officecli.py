"""OfficeCLI 的 Python 封装：subprocess 调用 + 结构化工具集。

经实测验证的关键点（officecli v1.0.138）:
    - subprocess 必须用 encoding='utf-8'，且 JSON 通过 --commands argv 传递
      （stdin 传中文 JSON 会被误解析；shell echo 也会转码）。
    - 默认 create 的 docx 只有 Normal 样式，Heading1/Title/ListBullet 不存在，
      因此本封装层用【显式格式 props】(size/bold/color/font/listStyle) 而非命名样式，
      确保 Word 打开时格式真实生效。
    - 表格元素路径是 /body/tbl[N]/tr[N]/tc[N]（注意是 tbl/tr/tc，不是 table/row/cell）。
    - 设置 OFFICECLI_NO_AUTO_RESIDENT=1 避免后台常驻进程文件锁。

设计:
    - 文本一律通过 argv 参数数组传递，绝不拼 shell 字符串，杜绝注入。
    - 写入优先用 batch（--commands JSON），失败整体回滚；add_table 内部用 batch 写单元格。
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from dataclasses import dataclass, field
from typing import Any

from .config import settings


class OfficeCLIError(RuntimeError):
    """officecli 调用失败。"""

    def __init__(self, message: str, *, cmd: list[str] | None = None,
                 returncode: int | None = None, stderr: str | None = None) -> None:
        super().__init__(message)
        self.cmd = cmd
        self.returncode = returncode
        self.stderr = stderr


def resolve_bin() -> str:
    """按优先级解析 officecli 可执行文件路径。

    1. 环境变量 OFFICECLI_BIN
    2. 工程内 bin/officecli 或 bin/officecli.exe
    3. PATH 中的 officecli
    """
    explicit = settings.officecli_bin.strip()
    if explicit:
        p = os.path.normpath(explicit)
        if os.path.exists(p):
            return p
        found = shutil.which(explicit)
        if found:
            return found
        raise OfficeCLIError(f"OFFICECLI_BIN 指定的路径不存在: {explicit}")

    candidates = [
        settings.project_root / "bin" / "officecli.exe",
        settings.project_root / "bin" / "officecli",
    ]
    for c in candidates:
        if c.exists():
            return str(c)

    found = shutil.which("officecli") or shutil.which("officecli.exe")
    if found:
        return found

    raise OfficeCLIError(
        "找不到 officecli 可执行文件。请执行:\n"
        "    python scripts/fetch_officecli.py\n"
        "下载到工程内 bin/，或在 .env 中设置 OFFICECLI_BIN 指向已有二进制。",
    )


@dataclass
class _Runner:
    """底层 subprocess 执行器。"""

    bin_path: str = field(default_factory=resolve_bin)

    def run(self, args: list[str], *, json_output: bool = False,
            timeout: int | None = None) -> Any:
        """执行一条 officecli 命令。

        args: 不含可执行文件本身的参数列表，如 ["create", "a.docx"]
        json_output: 是否追加 --json 并解析返回值
        返回: json_output=True 时返回解析后的对象；否则返回 stdout 文本
        """
        cmd = [self.bin_path, *args]
        if json_output and "--json" not in args:
            cmd.append("--json")

        env = dict(os.environ)
        # 关闭常驻模式，避免文件锁/后台进程干扰（每条命令独立 open/save）
        env.setdefault("OFFICECLI_NO_AUTO_RESIDENT", "1")

        try:
            proc = subprocess.run(
                cmd,
                env=env,
                capture_output=True,
                text=True,
                encoding="utf-8",  # 强制 UTF-8，避免 Windows GBK 乱码
                timeout=timeout or settings.officecli_timeout,
                check=False,
            )
        except FileNotFoundError as e:
            raise OfficeCLIError(f"无法执行 officecli: {e}", cmd=cmd) from e
        except subprocess.TimeoutExpired as e:
            raise OfficeCLIError(
                f"officecli 命令超时（{e.timeout}s）: {' '.join(args[:3])}...",
                cmd=cmd,
            ) from e

        if proc.returncode != 0:
            raise OfficeCLIError(
                f"officecli 返回非零状态 {proc.returncode}: {' '.join(args[:4])}\n"
                f"stderr: {proc.stderr.strip()}",
                cmd=cmd,
                returncode=proc.returncode,
                stderr=proc.stderr.strip(),
            )

        stdout = proc.stdout
        if json_output:
            stripped = stdout.strip()
            if not stripped:
                return None
            try:
                return json.loads(stripped)
            except json.JSONDecodeError:
                return stdout
        return stdout


# 模块级单例（延迟初始化，避免 import 时就要求二进制存在）
_runner: _Runner | None = None


def get_runner() -> _Runner:
    global _runner
    if _runner is None:
        _runner = _Runner()
    return _runner


def reset_runner(bin_path: str | None = None) -> None:
    """重置 runner，供测试或显式指定路径时使用。"""
    global _runner
    _runner = _Runner(bin_path=bin_path) if bin_path else _Runner()


# ============================================================
# 受限逃生口：白名单 raw
# ============================================================
_RAW_WHITELIST = {
    "create", "add", "set", "get", "query", "view",
    "close", "open", "save", "validate", "batch",
}


def raw(command_args: list[str], *, json_output: bool = False) -> Any:
    """受限逃生口：允许调用白名单内的子命令。

    command_args: 不含可执行文件的完整参数，如 ["view", "a.docx", "outline"]
    第一个元素必须是白名单中的子命令名。
    """
    if not command_args:
        raise OfficeCLIError("raw 调用不能为空")
    sub = command_args[0]
    if sub not in _RAW_WHITELIST:
        raise OfficeCLIError(
            f"raw 子命令 '{sub}' 不在白名单内，允许: {sorted(_RAW_WHITELIST)}",
        )
    return get_runner().run(command_args, json_output=json_output)


# ============================================================
# 面向文档的高层工具：DocTool
# ============================================================
# 标题字号映射（Word 约定，H1 最大），用显式 size 替代命名样式
_HEADING_SIZE_PT = {1: 22, 2: 18, 3: 15, 4: 13.5, 5: 12, 6: 11, 7: 11, 8: 11, 9: 11}
_TITLE_SIZE_PT = 26


@dataclass
class DocTool:
    """绑定单个文档路径的结构化操作工具。

    供 LangGraph 节点直接调用，语义清晰，内部翻译为 officecli 命令。
    所有格式用显式 props（size/bold/color/listStyle）确保 Word 真实生效，
    不依赖默认 docx 缺失的命名样式。
    """

    doc_path: str

    @property
    def runner(self) -> _Runner:
        return get_runner()

    # ---- 生命周期 ----
    def create(self) -> str:
        """创建空文档（若已存在会被覆盖）。"""
        return self.runner.run(["create", self.doc_path, "--force"])

    def close(self) -> str:
        """刷盘并释放（关闭常驻会话）。"""
        return self.runner.run(["close", self.doc_path])

    # ---- 写入 ----
    def add_title(self, text: str) -> str:
        """加文档主标题。显式大字号 + 加粗，居中。"""
        return self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", f"text={text}",
            "--prop", f"size={_TITLE_SIZE_PT}",
            "--prop", "bold=true",
            "--prop", "align=center",
        ])

    def add_heading(self, text: str, level: int = 1) -> str:
        """加章节标题。level 1-9，字号递减、加粗。"""
        level = max(1, min(9, int(level)))
        size = _HEADING_SIZE_PT[level]
        return self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", f"text={text}",
            "--prop", f"size={size}",
            "--prop", "bold=true",
            "--prop", "spaceBefore=12pt",
            "--prop", "spaceAfter=6pt",
        ])

    def add_paragraph(self, text: str, *, bold: bool = False,
                      italic: bool = False, size: float | None = None) -> str:
        """加正文段落。"""
        props = [f"text={text}"]
        if bold:
            props.append("bold=true")
        if italic:
            props.append("italic=true")
        if size is not None:
            props.append(f"size={size}")
        args = ["add", self.doc_path, "/body", "--type", "paragraph"]
        for p in props:
            args += ["--prop", p]
        return self.runner.run(args)

    def add_list_item(self, text: str, *, ordered: bool = False) -> str:
        """加列表项。用 listStyle=bullet/ordered（实测有效）。"""
        style = "ordered" if ordered else "bullet"
        return self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", f"text={text}",
            "--prop", f"listStyle={style}",
        ])

    def add_table(self, data: list[list[str]], *, has_header: bool = True) -> Any:
        """加表格。data 是二维字符串数组。

        实现: add table 建空表 → 查询新表真实索引 → batch 逐单元格写入。

        为什么不用 --prop data 一步创建: officecli 的 data CSV 格式【不支持
        引号转义】，单元格内含逗号/分号会被拆成多列（如"50,000"→两列），
        导致列数错乱。逐单元格 batch 写入对任意内容都安全。

        路径定位: 多表场景下新表索引是递增的（tbl[1]/tbl[2]...），不能写死。
        每次创建后用 _last_table_index() 查询 body 下最后一个 tbl 的索引，
        确保写入正确的表。
        """
        if not data or not data[0]:
            raise OfficeCLIError("表格数据为空")

        rows = len(data)
        cols = max(len(row) for row in data)
        # 补齐每行长度一致（防 LLM 给出参差不齐的数据）
        norm = [list(row) + [""] * (cols - len(row)) for row in data]

        # 建表并从输出直接解析新表索引（比 get 查询更可靠，不受图片等结构干扰）
        add_output = self.runner.run([
            "add", self.doc_path, "/body", "--type", "table",
            "--prop", f"rows={rows}",
            "--prop", f"cols={cols}",
        ])
        tbl_index = self._parse_tbl_index(add_output) or self._last_table_index()

        ops = self._build_table_ops(norm, tbl_index, has_header)
        try:
            self.batch(ops)
        except OfficeCLIError:
            # 重试：重新查询索引（兜底），重建 ops
            tbl_index = self._last_table_index()
            ops2 = self._build_table_ops(norm, tbl_index, has_header)
            try:
                self.batch(ops2)
            except OfficeCLIError:
                # 仍失败：逐单元格写（非原子，但保证尽量写入）
                for op in ops2:
                    try:
                        self.batch([op])
                    except OfficeCLIError:
                        pass  # 单个单元格失败不阻断其余

        return f"已添加 {rows} 行 × {cols} 列的表格"

    @staticmethod
    def _parse_tbl_index(output: str) -> int:
        """从 'add table' 的输出 'Added table at /body/tbl[N]' 解析索引 N。
        解析失败返回 0。"""
        import re
        m = re.search(r"tbl\[(\d+)\]", output or "")
        return int(m.group(1)) if m else 0

    @staticmethod
    def _build_table_ops(norm: list[list[str]], tbl_index: int,
                         has_header: bool) -> list[dict]:
        """构造写表格单元格的 batch ops。tbl_index 1-based。"""
        if tbl_index <= 0:
            tbl_index = 1
        ops: list[dict] = []
        for r, row in enumerate(norm):
            for c, cell in enumerate(row):
                props: dict[str, str] = {"text": str(cell)}
                if has_header and r == 0:
                    props["bold"] = "true"
                ops.append({
                    "command": "set",
                    "path": f"/body/tbl[{tbl_index}]/tr[{r + 1}]/tc[{c + 1}]",
                    "props": props,
                })
        return ops

    def _last_table_index(self) -> int:
        """查询 /body 下最后一个 table 的索引（1-based）。无表返回 0。"""
        try:
            data = self.runner.run(
                ["get", self.doc_path, "/body", "--depth", "1"],
                json_output=True,
            )
            children = (
                data.get("data", {}).get("results", [{}])[0].get("children", [])
                if isinstance(data, dict) else []
            )
            tbl_indices = [
                int(p.split("tbl[")[1].rstrip("]"))
                for c in children
                if (p := c.get("path", "")) and "tbl[" in p
            ]
            return max(tbl_indices) if tbl_indices else 0
        except Exception:  # noqa: BLE001
            return 0

    def batch(self, ops: list[dict]) -> Any:
        """批量原子操作。ops 见 officecli batch 文档。

        JSON 通过 --commands argv 传递（实测：stdin 传中文 JSON 会乱码/失败，
        argv 传递稳定且 UTF-8 round-trip 正确）。
        """
        payload = json.dumps(ops, ensure_ascii=False)
        return self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )

    def add_image(self, src: str, *, width: str = "8cm",
                  alt: str = "", caption: str = "") -> str:
        """插入图片到文档末尾。

        实现: picture 元素的 parent 必须是 paragraph（不是 body），所以
        先加一个段落承载图片，再把图片插入该段落。

        参数:
            src: 图片来源。本地文件路径 / URL / data URI 均可（officecli 支持）。
            width: 显示宽度（如 '8cm'/'400px'/'3in'）。高度按比例。
            alt: 图片替代文本（无障碍 + 图片加载失败时显示）。
            caption: 可选图注。非空时在图片下方加一段居中的图注文字。
        """
        # 1) 先加一个空段落作为图片的 parent
        self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
        ])
        # 2) 定位刚加的段落（body 下最后一个 p）
        p_index = self._last_paragraph_index()
        if p_index <= 0:
            p_index = 1
        # 3) 把图片插入该段落
        props = [f"src={src}", f"width={width}"]
        if alt:
            props.append(f"alt={alt}")
        args = ["add", self.doc_path, f"/body/p[{p_index}]", "--type", "picture"]
        for p in props:
            args += ["--prop", p]
        self.runner.run(args)

        # 4) 可选图注
        if caption:
            self.add_paragraph(caption, italic=True)  # 简化：图注作为斜体段落

        return f"已插入图片（{width}）" + (f"，图注: {caption}" if caption else "")

    def _last_paragraph_index(self) -> int:
        """查询 /body 下最后一个 paragraph 的索引（1-based）。无段落返回 0。"""
        try:
            data = self.runner.run(
                ["get", self.doc_path, "/body", "--depth", "1"],
                json_output=True,
            )
            children = (
                data.get("data", {}).get("results", [{}])[0].get("children", [])
                if isinstance(data, dict) else []
            )
            p_indices = [
                int(p.split("p[")[1].rstrip("]"))
                for c in children
                if (p := c.get("path", "")) and "/p[" in p
            ]
            return max(p_indices) if p_indices else 0
        except Exception:  # noqa: BLE001
            return 0

    # ---- 读取 ----
    def view_outline(self) -> str:
        """读文档大纲。"""
        return self.runner.run(["view", self.doc_path, "outline"])

    def view_text(self) -> str:
        """读文档纯文本（含路径标注，便于 review 节点定位）。"""
        return self.runner.run(["view", self.doc_path, "text"])

    def view_stats(self) -> Any:
        """读文档统计信息（json）。"""
        return self.runner.run(["view", self.doc_path, "stats"], json_output=True)

    def validate(self) -> str:
        """校验文档。"""
        return self.runner.run(["validate", self.doc_path])

    # ---- 目录 ----
    def add_toc(self, *, levels: str = "1-3", title: str = "",
                hyperlinks: bool = True, page_numbers: bool = True) -> str:
        """插入目录（Table of Contents）。

        参数:
            levels: 收录的标题层级范围，如 '1-3'。
            title: 目录标题（如"目录"）。留空则无标题。
            hyperlinks: 条目是否可点击跳转。
            page_numbers: 是否显示页码。

        注意: TOC 是 Word 域，真实条目需 Word 打开时刷新（或预先渲染）。
        会同时设 updateFields=true 让 Word 下次打开时自动重建条目。
        """
        args = [
            "add", self.doc_path, "/", "--type", "toc",
            "--prop", f"levels={levels}",
        ]
        if title:
            args += ["--prop", f"title={title}"]
        if hyperlinks:
            args += ["--prop", "hyperlinks=true"]
        if page_numbers:
            args += ["--prop", "pageNumbers=true"]
        result = self.runner.run(args)
        # 让 Word 打开时自动刷新 TOC 条目
        try:
            self.runner.run(["set", self.doc_path, "/", "--prop", "updateFields=true"])
        except OfficeCLIError:
            pass
        return result

    # ---- 页眉 / 页脚 ----
    def add_header(self, text: str = "", *, field: str = "",
                   align: str = "", kind: str = "default") -> str:
        """加页眉。

        参数:
            text: 页眉文字（与 field 二选一或组合）。
            field: 域类型，如 'page'(页码)/'numpages'(总页数)/'date'/'author'/'title'。
            align: 对齐 left/center/right。
            kind: 'default'(所有页)/'first'(首页不同)/'even'(偶数页不同)。
        """
        args = ["add", self.doc_path, "/", "--type", "header",
                "--prop", f"type={kind}"]
        if text:
            args += ["--prop", f"text={text}"]
        if field:
            args += ["--prop", f"field={field}"]
        if align:
            args += ["--prop", f"align={align}"]
        return self.runner.run(args)

    def add_footer(self, text: str = "", *, field: str = "",
                   align: str = "center", kind: str = "default") -> str:
        """加页脚。参数同 add_header，默认居中。"""
        args = ["add", self.doc_path, "/", "--type", "footer",
                "--prop", f"type={kind}"]
        if text:
            args += ["--prop", f"text={text}"]
        if field:
            args += ["--prop", f"field={field}"]
        if align:
            args += ["--prop", f"align={align}"]
        return self.runner.run(args)

    # ---- 超链接 ----
    def add_hyperlink(self, text: str, url: str = "", *, anchor: str = "",
                      tooltip: str = "") -> str:
        """加一个超链接段落（外部 URL 或内部书签）。

        参数:
            text: 显示文字。
            url: 外部链接（如 'https://example.com'）。与 anchor 二选一。
            anchor: 内部书签名（需先用 add_bookmark 创建）。
            tooltip: 鼠标悬停提示。

        实现: 先加一个承载段落，再在该段落下加 hyperlink 元素。
        """
        # 加空段落
        self.runner.run(["add", self.doc_path, "/body", "--type", "paragraph"])
        p_index = self._last_paragraph_index() or 1
        args = ["add", self.doc_path, f"/body/p[{p_index}]", "--type", "hyperlink",
                "--prop", f"text={text}"]
        if url:
            args += ["--prop", f"url={url}"]
        elif anchor:
            args += ["--prop", f"anchor={anchor}"]
        if tooltip:
            args += ["--prop", f"tooltip={tooltip}"]
        return self.runner.run(args)

    # ---- 图表 ----
    def add_chart(self, chart_type: str, data: str, *,
                  categories: str = "", title: str = "",
                  width: str = "15cm", height: str = "8cm",
                  legend: str = "bottom") -> str:
        """加嵌入式图表（Word 图表自带数据，不引用 Excel）。

        参数:
            chart_type: 'column'/'bar'/'line'/'pie'/'doughnut'/'area'/'scatter'。
            data: 内联数据，格式 '系列名:值,值,值'，多系列分号分隔。
                  例: 'Sales:10,20,30;Cost:5,8,12'
            categories: 分类标签，逗号分隔，如 'Q1,Q2,Q3,Q4'。
            title: 图表标题。
            width/height: 图表尺寸。
            legend: 图例位置 none/top/bottom/left/right。
        """
        # 图表 parent 必须是段落
        self.runner.run(["add", self.doc_path, "/body", "--type", "paragraph"])
        p_index = self._last_paragraph_index() or 1
        args = [
            "add", self.doc_path, f"/body/p[{p_index}]", "--type", "chart",
            "--prop", f"chartType={chart_type}",
            "--prop", f"data={data}",
            "--prop", f"width={width}", "--prop", f"height={height}",
            "--prop", f"legend={legend}",
        ]
        if categories:
            args += ["--prop", f"categories={categories}"]
        if title:
            args += ["--prop", f"title={title}"]
        return self.runner.run(args)

    # ---- 分节 ----
    def add_section(self, *, section_type: str = "nextPage",
                    orientation: str = "",
                    columns: int = 1) -> str:
        """插入分节符。

        参数:
            section_type: 'nextPage'(下一页)/'continuous'(连续)/'evenPage'/'oddPage'。
            orientation: 'portrait'/'landscape'（留空不变）。
            columns: 分栏数（1=不分栏）。
        """
        args = ["add", self.doc_path, "/", "--type", "section",
                "--prop", f"type={section_type}"]
        if orientation:
            args += ["--prop", f"orientation={orientation}"]
        if columns > 1:
            args += ["--prop", f"columns={columns}"]
        return self.runner.run(args)

    # ---- 文档属性 ----
    def set_doc_properties(self, *, title: str = "", author: str = "",
                           subject: str = "", keywords: str = "",
                           description: str = "") -> str:
        """设置文档核心属性（标题/作者/主题/关键词/摘要）。

        这些属性显示在 Word 的"文件 → 信息"和文件资源管理器属性里。
        """
        args = ["set", self.doc_path, "/"]
        if title:
            args += ["--prop", f"title={title}"]
        if author:
            args += ["--prop", f"author={author}"]
        if subject:
            args += ["--prop", f"subject={subject}"]
        if keywords:
            args += ["--prop", f"keywords={keywords}"]
        if description:
            args += ["--prop", f"description={description}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_doc_properties: 至少传一个属性")
        return self.runner.run(args)

    # ---- 按命名样式加段落 ----
    def add_styled_paragraph(self, text: str, style: str) -> str:
        """按命名样式加段落（如 'Heading1'/'Title'/'Quote'/自定义样式）。

        与 add_heading 的区别: add_heading 用显式 size/bold（默认文档也能用），
        本方法用真实命名样式（需文档已有该样式，或先用 add_style 创建）。
        用真实 Heading 样式才能让 TOC 自动收录。
        """
        return self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", f"text={text}",
            "--prop", f"style={style}",
        ])

    def add_style(self, style_id: str, name: str, *,
                  style_type: str = "paragraph", based_on: str = "Normal",
                  size: float | None = None, bold: bool = False,
                  color: str = "", outline_level: int | None = None) -> str:
        """创建命名样式。

        参数:
            style_id: 样式 ID（如 'MyHeading'）。
            name: 显示名。
            style_type: 'paragraph'/'character'/'table'。
            based_on: 继承自哪个样式（默认 Normal）。
            outline_level: 大纲级别 0-9（0=一级标题，能让 TOC 收录）。
        """
        args = [
            "add", self.doc_path, "/styles", "--type", "style",
            "--prop", f"id={style_id}",
            "--prop", f"name={name}",
            "--prop", f"type={style_type}",
            "--prop", f"basedOn={based_on}",
        ]
        if size is not None:
            args += ["--prop", f"size={size}"]
        if bold:
            args += ["--prop", "bold=true"]
        if color:
            args += ["--prop", f"color={color}"]
        if outline_level is not None:
            args += ["--prop", f"outlineLvl={outline_level}"]
        return self.runner.run(args)


# ============================================================
# 面向 Excel 的高层工具：ExcelTool
# ============================================================
# 单元格路径 /<sheetName>/<A1Ref>（如 /Sheet1/A1、/Data/B2:C3）。
# 写单元格一律用 `set`（officecli 的 set 会自动创建不存在的单元格，
# add 则不会）。批量写用 batch + 逐 cell 的 set op，避免 officecli
# CSV data 的引号转义陷阱（与 DocTool.add_table 同一思路）。


def _col_to_letter(n: int) -> str:
    """1-based 列号 → Excel 列字母（1→A, 27→AA）。"""
    n = max(1, int(n))
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def _ref_at(row: int, col: int) -> str:
    """1-based (行, 列) → A1 引用（如 (2,3)→'C2'）。"""
    return f"{_col_to_letter(col)}{int(row)}"


@dataclass
class ExcelTool:
    """绑定单个 .xlsx 路径的 Excel 操作工具。

    供 LangGraph 节点直接调用，内部翻译为 officecli 命令。
    单元格值用 `set` 写入（auto-vivify）；批量写用 batch 逐 cell。
    """

    doc_path: str

    @property
    def runner(self) -> _Runner:
        return get_runner()

    # ---- 生命周期 ----
    def create(self) -> str:
        """创建空 xlsx（若已存在会被覆盖）。"""
        return self.runner.run(["create", self.doc_path, "--force"])

    def close(self) -> str:
        """刷盘并释放。"""
        return self.runner.run(["close", self.doc_path])

    # ---- 工作表 ----
    def add_sheet(self, name: str, *, tab_color: str = "") -> str:
        """加工作表。
        参数:
            name: 工作表名（如 '销售数据'）。
            tab_color: 可选标签色（如 '4472C4'）。
        """
        args = ["add", self.doc_path, "/", "--type", "sheet",
                "--prop", f"name={name}"]
        if tab_color:
            args += ["--prop", f"tabColor={tab_color}"]
        return self.runner.run(args)

    def rename_sheet(self, old: str, new: str) -> str:
        return self.runner.run([
            "set", self.doc_path, f"/{old}", "--prop", f"name={new}",
        ])

    def set_sheet_color(self, name: str, color: str) -> str:
        return self.runner.run([
            "set", self.doc_path, f"/{name}", "--prop", f"tabColor={color}",
        ])

    def set_doc_properties(self, *, title: str = "", author: str = "",
                           subject: str = "", keywords: str = "",
                           description: str = "") -> str:
        """设置工作簿核心属性（标题/作者/主题/关键词）。"""
        args = ["set", self.doc_path, "/"]
        if title:
            args += ["--prop", f"title={title}"]
        if author:
            args += ["--prop", f"author={author}"]
        if subject:
            args += ["--prop", f"subject={subject}"]
        if keywords:
            args += ["--prop", f"keywords={keywords}"]
        if description:
            args += ["--prop", f"description={description}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_doc_properties: 至少传一个属性")
        return self.runner.run(args)

    # ---- 单元格写入 ----
    def set_cell(self, sheet: str, ref: str, value: Any, *,
                 bold: bool = False, italic: bool = False,
                 fill: str = "", font_color: str = "",
                 size: float | None = None, align: str = "",
                 number_format: str = "") -> str:
        """写单个单元格。ref 如 'A1'。

        value 会转成字符串。数字字符串如 '01234' 需按文本存时传 number_format='@'。
        """
        props: list[str] = [f"value={value}"]
        if bold:
            props.append("bold=true")
        if italic:
            props.append("italic=true")
        if fill:
            props.append(f"fill={fill}")
        if font_color:
            props.append(f"font.color={font_color}")
        if size is not None:
            props.append(f"size={size}")
        if align:
            props.append(f"alignment.horizontal={align}")
        if number_format:
            props.append(f"numberformat={number_format}")
        args = ["set", self.doc_path, f"/{sheet}/{ref}"]
        for p in props:
            args += ["--prop", p]
        return self.runner.run(args)

    def set_cells(self, sheet: str, rows: list[list], *,
                  start_ref: str = "A1", has_header: bool = False) -> str:
        """批量写二维数组到工作表。

        参数:
            sheet: 目标工作表名。
            rows: 二维数组，外层=行、内层=单元格（数字/字符串均可）。
            start_ref: 起始单元格（默认 'A1'）。内部按行列递增算出每个 cell 的 ref。
            has_header: 第一行是否作为表头（加粗 + 可选底色）。

        实现: 把 rows 展开成 batch ops，逐 cell 用 `set` 写入。
        避免了 officecli CSV data 不支持引号转义、逗号/换行会拆列的问题。
        """
        if not rows:
            raise OfficeCLIError("set_cells: rows 为空")
        import re
        m = re.match(r"^([A-Za-z]+)(\d+)$", start_ref.strip())
        if not m:
            raise OfficeCLIError(f"set_cells: 非法 start_ref '{start_ref}'")
        col_letters, start_row = m.group(1), int(m.group(2))
        # 列字母 → 1-based 列号
        start_col = 0
        for ch in col_letters.upper():
            start_col = start_col * 26 + (ord(ch) - 64)

        # 补齐每行长度一致
        cols = max(len(r) for r in rows)
        ops: list[dict] = []
        for ri, row in enumerate(rows):
            for ci in range(cols):
                cell = row[ci] if ci < len(row) else ""
                cell_str = "" if cell is None else str(cell)
                props: dict[str, str] = {"value": cell_str}
                if has_header and ri == 0:
                    props["bold"] = "true"
                ref = _ref_at(start_row + ri, start_col + ci)
                ops.append({
                    "command": "set",
                    "path": f"/{sheet}/{ref}",
                    "props": props,
                })
        try:
            self.batch(ops)
        except OfficeCLIError:
            # 兜底：逐单元格写，单个失败不阻断其余
            for op in ops:
                try:
                    self.batch([op])
                except OfficeCLIError:
                    pass
        return f"已写入 {len(rows)} 行 × {cols} 列（起始 {start_ref}）"

    def set_formula(self, sheet: str, ref: str, formula: str) -> str:
        """写公式（不带前导 =）。如 formula='SUM(A1:A10)'。"""
        return self.runner.run([
            "set", self.doc_path, f"/{sheet}/{ref}",
            "--prop", f"formula={formula}",
        ])

    # ---- 行列格式 ----
    def set_column_width(self, sheet: str, col: str, width: float) -> str:
        """col 是列字母（如 'A'）或数字；width 是字符单位。"""
        return self.runner.run([
            "set", self.doc_path, f"/{sheet}/col[{col}]",
            "--prop", f"width={width}",
        ])

    def set_row_height(self, sheet: str, row: int, height: float) -> str:
        """row 是 1-based 行号；height 是磅。"""
        return self.runner.run([
            "set", self.doc_path, f"/{sheet}/row[{row}]",
            "--prop", f"height={height}",
        ])

    def autofit_column(self, sheet: str, col: str) -> str:
        return self.runner.run([
            "set", self.doc_path, f"/{sheet}/col[{col}]",
            "--prop", "autofit=true",
        ])

    def merge_cells(self, sheet: str, cell_range: str) -> str:
        """合并单元格。cell_range 如 'A1:C3'，锚点必须是左上格。"""
        anchor = cell_range.split(":")[0]
        return self.runner.run([
            "set", self.doc_path, f"/{sheet}/{anchor}",
            "--prop", f"merge={cell_range}",
        ])

    # ---- 图表 ----
    def add_chart(self, sheet: str, chart_type: str, data_range: str, *,
                  categories: str = "", title: str = "",
                  x: str = "2cm", y: str = "10cm",
                  width: str = "15cm", height: str = "8cm",
                  legend: str = "bottom") -> str:
        """加图表（基于工作表已有数据）。

        参数:
            sheet: 图表所在工作表名。
            chart_type: 图表类型，如 'column'/'bar'/'line'/'pie'/'doughnut'/'area'/'scatter'。
            data_range: 数据源区域，如 'Sheet1!B2:C5'（注意要带工作表名前缀）。
                        默认首列当分类轴；若传 categories 则每列都是一个系列。
            categories: 可选分类轴区域或逗号分隔标签（如 'Sheet1!A2:A5' 或 'Q1,Q2,Q3'）。
            title: 图表标题。
            x/y/width/height: 图表在工作表上的位置和尺寸。
            legend: 图例位置（'none'/'top'/'bottom'/'left'/'right'）。
        """
        args = [
            "add", self.doc_path, f"/{sheet}", "--type", "chart",
            "--prop", f"chartType={chart_type}",
            "--prop", f"dataRange={data_range}",
            "--prop", f"x={x}", "--prop", f"y={y}",
            "--prop", f"width={width}", "--prop", f"height={height}",
            "--prop", f"legend={legend}",
        ]
        if categories:
            args += ["--prop", f"categories={categories}"]
        if title:
            args += ["--prop", f"title={title}"]
        return self.runner.run(args)

    # ---- 通用 ----
    def batch(self, ops: list[dict]) -> Any:
        """批量原子操作。复用 DocTool 同款实现（argv 传 JSON）。"""
        payload = json.dumps(ops, ensure_ascii=False)
        return self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )

    # ---- 排序 / 筛选 ----
    def sort(self, sheet: str, keys: str, *, has_header: bool = True) -> str:
        """排序工作表数据。

        参数:
            sheet: 工作表名。
            keys: 排序键，格式 '列字母 [asc|desc]'，多键逗号分隔。
                  例: 'B desc' / 'A asc, B desc' / 'C'
            has_header: 首行是否为表头（不参与排序）。
        """
        args = ["set", self.doc_path, f"/{sheet}", "--prop", f"sort={keys}"]
        if has_header:
            args += ["--prop", "sortHeader=true"]
        return self.runner.run(args)

    def set_autofilter(self, sheet: str, cell_range: str = "") -> str:
        """开启自动筛选。cell_range 留空则用已用区域。

        参数:
            sheet: 工作表名。
            cell_range: 筛选区域，如 'A1:D10'。留空=自动识别已用区域。
        """
        val = cell_range if cell_range else "true"
        return self.runner.run([
            "set", self.doc_path, f"/{sheet}", "--prop", f"autoFilter={val}",
        ])

    # ---- 条件格式 ----
    def add_conditional_format(self, sheet: str, cf_type: str,
                               cell_range: str, **props) -> str:
        """加条件格式规则。

        参数:
            sheet: 工作表名。
            cf_type: 规则类型。常用:
                'cellIs'           单元格值比较（需 operator + value + fill）
                'colorScale'       3色渐变（需 minColor/midColor/maxColor）
                'dataBar'          数据条（需 color）
                'topN'             前 N 项（需 rank）
                'duplicateValues'  重复值高亮（需 fill）
                'iconSet'          图标集（需 iconset）
            cell_range: 应用区域，如 'A1:A10'。
            **props: 其余属性，如 operator='greaterThan' value='50' fill='FFFF00'。
                    直接用关键字参数传，内部转 --prop key=val。

        示例:
            add_conditional_format('S', 'cellIs', 'A1:A10',
                                   operator='greaterThan', value='50', fill='FFFF00')
        """
        args = [
            "add", self.doc_path, f"/{sheet}", "--type", cf_type,
            "--prop", f"sqref={cell_range}",
        ]
        for k, v in props.items():
            args += ["--prop", f"{k}={v}"]
        return self.runner.run(args)

    # ---- 透视表 ----
    def add_pivot_table(self, sheet: str, source: str, *,
                        rows: str = "", cols: str = "",
                        values: str = "", filters: str = "",
                        position: str = "", name: str = "",
                        layout: str = "") -> str:
        """加数据透视表。

        参数:
            sheet: 透视表放在哪张工作表。
            source: 源数据区域，如 'Sheet1!A1:D100'（必须含表头行）。
            rows: 行字段，逗号分隔，如 'Region,Category'。
            cols: 列字段，如 'Quarter'。
            values: 值字段及聚合方式，格式 '字段:agg'，多个逗号分隔。
                    agg ∈ sum/avg/count/max/min。例: 'Sales:sum,Qty:count'。
            filters: 筛选页字段，如 'Year'。
            position: 透视表左上角位置，如 'H1'。留空自动放在源数据旁。
            name: 透视表名。
            layout: 布局 compact/outline/tabular。
        """
        args = [
            "add", self.doc_path, f"/{sheet}", "--type", "pivottable",
            "--prop", f"source={source}",
        ]
        if rows:
            args += ["--prop", f"rows={rows}"]
        if cols:
            args += ["--prop", f"cols={cols}"]
        if values:
            args += ["--prop", f"values={values}"]
        if filters:
            args += ["--prop", f"filters={filters}"]
        if position:
            args += ["--prop", f"position={position}"]
        if name:
            args += ["--prop", f"name={name}"]
        if layout:
            args += ["--prop", f"layout={layout}"]
        return self.runner.run(args)

    # ---- ListObject 表格 ----
    def add_list_table(self, sheet: str, cell_range: str, *,
                       name: str = "", style: str = "medium2",
                       total_row: bool = False) -> str:
        """把单元格区域转成真正的 Excel 表格（带样式、筛选按钮、结构化引用）。

        参数:
            sheet: 工作表名。
            cell_range: 区域，如 'A1:C10'（首行需是表头）。
            name: 表格名。
            style: 表样式。常用 'medium1'~'medium4', 'light1'~'light3', 'dark1'~'dark2'。
            total_row: 是否显示汇总行。
        """
        args = [
            "add", self.doc_path, f"/{sheet}", "--type", "table",
            "--prop", f"ref={cell_range}",
            "--prop", f"style={style}",
        ]
        if name:
            args += ["--prop", f"name={name}"]
        if total_row:
            args += ["--prop", "totalRow=true"]
        return self.runner.run(args)

    # ---- 命名区域 ----
    def add_named_range(self, name: str, refers_to: str, *,
                        scope: str = "workbook", comment: str = "") -> str:
        """定义命名区域（供公式用名称引用）。

        参数:
            name: 名称（如 'Revenue'）。
            refers_to: 引用，【不带前导 =】，如 'Sheet1!$A$1:$C$10'。
            scope: 作用域，'workbook' 或工作表名。
            comment: 备注（Name Manager 里显示）。
        """
        args = [
            "add", self.doc_path, "/", "--type", "namedrange",
            "--prop", f"name={name}",
            "--prop", f"ref={refers_to}",
            "--prop", f"scope={scope}",
        ]
        if comment:
            args += ["--prop", f"comment={comment}"]
        return self.runner.run(args)

    # ---- 数据验证 ----
    def add_validation(self, sheet: str, cell_range: str, val_type: str,
                       *, formula1: str = "", formula2: str = "",
                       operator: str = "", in_cell_dropdown: bool = True,
                       prompt: str = "", error_msg: str = "") -> str:
        """加数据验证（下拉列表 / 数值范围等）。

        参数:
            sheet: 工作表名。
            cell_range: 应用区域，如 'B1:B10'。
            val_type: 验证类型 'list'(下拉)/'whole'(整数)/'decimal'(小数)/
                      'date'/'time'/'textLength'/'custom'。
            formula1: 主值。list 时是逗号分隔的选项（如 'Yes,No,Maybe'）；
                      whole/decimal 时是下界或比较值。
            formula2: 上界，仅 between/notBetween 用。
            operator: 比较运算符 between/notBetween/equal/notEqual/
                      greaterThan/lessThan 等（list 不用）。
            in_cell_dropdown: list 类型时显示下拉箭头。
            prompt: 输入提示。
            error_msg: 输入非法时的错误提示。
        """
        args = [
            "add", self.doc_path, f"/{sheet}", "--type", "validation",
            "--prop", f"ref={cell_range}",
            "--prop", f"type={val_type}",
        ]
        if formula1:
            args += ["--prop", f"formula1={formula1}"]
        if formula2:
            args += ["--prop", f"formula2={formula2}"]
        if operator:
            args += ["--prop", f"operator={operator}"]
        if val_type == "list":
            args += ["--prop", f"inCellDropdown={'true' if in_cell_dropdown else 'false'}"]
        if prompt:
            args += ["--prop", f"prompt={prompt}"]
        if error_msg:
            args += ["--prop", f"error={error_msg}", "--prop", "errorTitle=输入错误"]
        return self.runner.run(args)

    # ---- 读取单元格（含公式计算值）----
    def get_cell(self, sheet: str, ref: str) -> dict:
        """读取单元格的值、公式、缓存值/计算值（json）。

        返回 dict 含: value, formula(若有), cachedValue(上次Excel算的),
        computedValue(officecli算的), evaluated。
        用于验证公式结果是否正确。
        """
        return self.runner.run(
            ["get", self.doc_path, f"/{sheet}/{ref}", "--json"],
            json_output=True,
        )

    # ---- 读取 ----
    def view_text(self) -> str:
        """读工作表纯文本（输出 A1=value\\tB1=value 形式，便于 LLM 自查）。"""
        return self.runner.run(["view", self.doc_path, "text"])

    def view_stats(self) -> Any:
        return self.runner.run(["view", self.doc_path, "stats"], json_output=True)

    def validate(self) -> str:
        return self.runner.run(["validate", self.doc_path])


# ============================================================
# 面向 PowerPoint 的高层工具：PptxTool
# ============================================================
# 幻灯片路径 /slide[N]（1-based）。shape 路径优先用 add 返回的
# /slide[N]/shape[@id=ID]（positional /shape[M] 会混入布局占位符）。
# 表格单元格路径是 /slide[N]/table[M]/tr[R]/tc[C]（OOXML 元素名 tr/tc）。


@dataclass
class PptxTool:
    """绑定单个 .pptx 路径的 PowerPoint 操作工具。

    供 LangGraph 节点直接调用，内部翻译为 officecli 命令。
    """

    doc_path: str

    @property
    def runner(self) -> _Runner:
        return get_runner()

    # ---- 生命周期 ----
    def create(self) -> str:
        """创建空 pptx（若已存在会被覆盖）。"""
        return self.runner.run(["create", self.doc_path, "--force"])

    def close(self) -> str:
        return self.runner.run(["close", self.doc_path])

    # ---- 幻灯片 ----
    def add_slide(self, *, title: str = "", text: str = "",
                  layout: str = "") -> str:
        """加一张幻灯片。

        参数:
            title: 标题文字（非空时自动建标题占位符 phType=title）。
            text: 正文文字（非空时自动建正文占位符 phType=body）。
            layout: 可选布局名（如 'Title Slide' / 'Title and Content' / 'Blank'）。
                    留空用默认布局。

        返回 officecli 的 add 输出（含新幻灯片路径）。
        """
        args = ["add", self.doc_path, "/", "--type", "slide"]
        if title:
            args += ["--prop", f"title={title}"]
        if text:
            args += ["--prop", f"text={text}"]
        if layout:
            args += ["--prop", f"layout={layout}"]
        return self.runner.run(args)

    def _last_slide_index(self) -> int:
        """查询最后一张幻灯片的索引（1-based）。无返回 0。"""
        try:
            data = self.runner.run(
                ["get", self.doc_path, "/", "--depth", "1"],
                json_output=True,
            )
            children = (
                data.get("data", {}).get("results", [{}])[0].get("children", [])
                if isinstance(data, dict) else []
            )
            idxs = [
                int(p.split("slide[")[1].rstrip("]"))
                for c in children
                if (p := c.get("path", "")) and "slide[" in p
            ]
            return max(idxs) if idxs else 0
        except Exception:  # noqa: BLE001
            return 0

    # ---- 文本框 / 形状 ----
    def add_textbox(self, slide_index: int, text: str, *,
                    x: str = "1cm", y: str = "2cm",
                    width: str = "22cm", height: str = "2cm",
                    size: float | None = None, bold: bool = False,
                    italic: bool = False, color: str = "",
                    fill: str = "", align: str = "left",
                    valign: str = "top") -> str:
        """在指定幻灯片上加文本框。

        参数:
            slide_index: 幻灯片 1-based 序号。
            text: 文本内容（可含换行）。
            x/y/width/height: 位置和尺寸（带单位，如 '2cm'）。
            size: 字号（pt）。
            bold/italic/color: 字体格式。
            fill: 文本框背景色（如 'FFFF00'）。
            align: 水平对齐 'left'/'center'/'right'/'justify'。
            valign: 垂直对齐 'top'/'center'/'bottom'。
        """
        args = [
            "add", self.doc_path, f"/slide[{slide_index}]", "--type", "textbox",
            "--prop", f"text={text}",
            "--prop", f"x={x}", "--prop", f"y={y}",
            "--prop", f"width={width}", "--prop", f"height={height}",
            "--prop", f"align={align}", "--prop", f"valign={valign}",
        ]
        if size is not None:
            args += ["--prop", f"size={size}"]
        if bold:
            args += ["--prop", "bold=true"]
        if italic:
            args += ["--prop", "italic=true"]
        if color:
            args += ["--prop", f"color={color}"]
        if fill:
            args += ["--prop", f"fill={fill}"]
        return self.runner.run(args)

    def add_shape(self, slide_index: int, text: str, *,
                  geometry: str = "rect",
                  x: str = "2cm", y: str = "2cm",
                  width: str = "4cm", height: str = "2cm",
                  fill: str = "", line: str = "",
                  size: float | None = None, bold: bool = False,
                  color: str = "", align: str = "center",
                  valign: str = "middle") -> str:
        """在指定幻灯片上加自选图形（矩形/椭圆/箭头等）+ 文字。

        参数:
            geometry: 形状预设，如 'rect'/'roundRect'/'ellipse'/'triangle'/
                      'diamond'/'rightArrow'/'star5'。
            line: 边框，格式 'color[:width[:style]]'（如 'FF0000:1.5:dash'）。
            其余参数同 add_textbox。
        """
        args = [
            "add", self.doc_path, f"/slide[{slide_index}]", "--type", "shape",
            "--prop", f"text={text}",
            "--prop", f"geometry={geometry}",
            "--prop", f"x={x}", "--prop", f"y={y}",
            "--prop", f"width={width}", "--prop", f"height={height}",
            "--prop", f"align={align}", "--prop", f"valign={valign}",
        ]
        if fill:
            args += ["--prop", f"fill={fill}"]
        if line:
            args += ["--prop", f"line={line}"]
        if size is not None:
            args += ["--prop", f"size={size}"]
        if bold:
            args += ["--prop", "bold=true"]
        if color:
            args += ["--prop", f"color={color}"]
        return self.runner.run(args)

    # ---- 图片 ----
    def add_image(self, slide_index: int, src: str, *,
                  x: str = "2cm", y: str = "2cm",
                  width: str = "10cm", height: str = "",
                  alt: str = "") -> str:
        """在指定幻灯片上插入图片。

        参数:
            slide_index: 幻灯片 1-based 序号。
            src: 图片来源（本地路径 / URL / data URI）。
            x/y/width/height: 位置和尺寸。height 留空则按图片比例。
            alt: 替代文本。
        """
        args = [
            "add", self.doc_path, f"/slide[{slide_index}]", "--type", "picture",
            "--prop", f"src={src}",
            "--prop", f"x={x}", "--prop", f"y={y}",
            "--prop", f"width={width}",
        ]
        if height:
            args += ["--prop", f"height={height}"]
        if alt:
            args += ["--prop", f"alt={alt}"]
        return self.runner.run(args)

    # ---- 表格 ----
    def add_table(self, slide_index: int, data: list[list], *,
                  has_header: bool = True,
                  x: str = "1cm", y: str = "4cm",
                  width: str = "22cm") -> str:
        """在指定幻灯片上加表格。

        参数:
            slide_index: 幻灯片 1-based 序号。
            data: 二维数组（外层=行、内层=单元格）。
            has_header: 第一行作为表头（加粗）。
            x/y/width: 表格位置和总宽。

        实现: add table 建空表（rows/cols）→ 查询新表在当前幻灯片的索引 →
        batch 逐单元格 set text（路径 /slide[N]/table[M]/tr[R]/tc[C]）。
        """
        if not data or not data[0]:
            raise OfficeCLIError("add_table: data 为空")

        rows = len(data)
        cols = max(len(r) for r in data)
        norm = [list(r) + [""] * (cols - len(r)) for r in data]

        self.runner.run([
            "add", self.doc_path, f"/slide[{slide_index}]", "--type", "table",
            "--prop", f"rows={rows}",
            "--prop", f"cols={cols}",
            "--prop", f"x={x}", "--prop", f"y={y}",
            "--prop", f"width={width}",
        ])
        tbl_index = self._last_table_index(slide_index)

        ops = self._build_table_ops(slide_index, tbl_index, norm, has_header)
        try:
            self.batch(ops)
        except OfficeCLIError:
            # 兜底：逐单元格
            for op in ops:
                try:
                    self.batch([op])
                except OfficeCLIError:
                    pass
        return f"已在幻灯片 {slide_index} 添加 {rows} 行 × {cols} 列的表格"

    def _last_table_index(self, slide_index: int) -> int:
        """查询指定幻灯片下最后一个 table 的索引（1-based）。无返回 0。"""
        try:
            data = self.runner.run(
                ["get", self.doc_path, f"/slide[{slide_index}]", "--depth", "1"],
                json_output=True,
            )
            children = (
                data.get("data", {}).get("results", [{}])[0].get("children", [])
                if isinstance(data, dict) else []
            )
            idxs = [
                int(p.split("table[")[1].rstrip("]"))
                for c in children
                if (p := c.get("path", "")) and "table[" in p
            ]
            return max(idxs) if idxs else 0
        except Exception:  # noqa: BLE001
            return 0

    @staticmethod
    def _build_table_ops(slide_index: int, tbl_index: int,
                         norm: list[list[str]], has_header: bool) -> list[dict]:
        """构造写 pptx 表格单元格的 batch ops。"""
        if tbl_index <= 0:
            tbl_index = 1
        ops: list[dict] = []
        for r, row in enumerate(norm):
            for c, cell in enumerate(row):
                props: dict[str, str] = {"text": str(cell)}
                if has_header and r == 0:
                    props["bold"] = "true"
                ops.append({
                    "command": "set",
                    "path": (f"/slide[{slide_index}]/table[{tbl_index}]"
                             f"/tr[{r + 1}]/tc[{c + 1}]"),
                    "props": props,
                })
        return ops

    # ---- 通用 ----
    def batch(self, ops: list[dict]) -> Any:
        """批量原子操作。复用 DocTool 同款实现（argv 传 JSON）。"""
        payload = json.dumps(ops, ensure_ascii=False)
        return self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )

    # ---- 幻灯片级元数据（与占位符模式兼容，不会重叠）----
    def set_transition(self, slide_index: int, transition: str, *,
                       advance_time_ms: int | None = None,
                       advance_on_click: bool = True) -> str:
        """设置幻灯片切换效果。

        参数:
            slide_index: 幻灯片 1-based 序号。
            transition: 切换效果。常用:
                'fade'/'none'    淡入/无
                'push'/'push-right'/'push-left'/'push-up'/'push-down'  推移
                'wipe'/'cover'/'split'/'zoom'/'cut'/'dissolve'
                'morph'          平滑（形状变形，需相邻页有同名形状）
                'circle'/'diamond'/'wedge'/'wheel'/'blinds'/'checker'
                可加速/方向后缀: 'push-right-1500'（推移右、1500ms）。
            advance_time_ms: 自动换页毫秒数（如 5000=5秒）。None=不自动换页。
            advance_on_click: 是否点击换页。
        """
        args = ["set", self.doc_path, f"/slide[{slide_index}]",
                "--prop", f"transition={transition}"]
        if advance_time_ms is not None:
            args += ["--prop", f"advanceTime={advance_time_ms}"]
        args += ["--prop", f"advanceClick={'true' if advance_on_click else 'false'}"]
        return self.runner.run(args)

    def set_notes(self, slide_index: int, text: str) -> str:
        """设置幻灯片的演讲者备注。

        参数:
            slide_index: 幻灯片 1-based 序号。
            text: 备注全文（可含换行）。
        """
        return self.runner.run([
            "set", self.doc_path, f"/slide[{slide_index}]",
            "--prop", f"notes={text}",
        ])

    def set_slide_hidden(self, slide_index: int, hidden: bool = True) -> str:
        """隐藏/显示幻灯片（放映时跳过）。"""
        return self.runner.run([
            "set", self.doc_path, f"/slide[{slide_index}]",
            "--prop", f"hidden={'true' if hidden else 'false'}",
        ])

    # ---- 演示文稿级（deck-wide）----
    def set_theme_colors(self, *, accent1: str = "", accent2: str = "",
                         accent3: str = "", accent4: str = "",
                         accent5: str = "", accent6: str = "",
                         hyperlink: str = "") -> str:
        """自定义主题颜色（影响全 deck 的强调色/超链接色）。

        参数: 6 个 accent 颜色 + hyperlink，6位十六进制（如 '4472C4'）。
        只传需要改的，其余保持不变。
        """
        args = ["set", self.doc_path, "/theme"]
        for k, v in [("accent1", accent1), ("accent2", accent2),
                     ("accent3", accent3), ("accent4", accent4),
                     ("accent5", accent5), ("accent6", accent6),
                     ("hyperlink", hyperlink)]:
            if v:
                args += ["--prop", f"{k}={v}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_theme_colors: 至少传一个颜色")
        return self.runner.run(args)

    def set_theme_fonts(self, heading_font: str = "",
                        body_font: str = "") -> str:
        """自定义主题字体（标题字体 / 正文字体，影响全 deck）。"""
        args = ["set", self.doc_path, "/theme"]
        if heading_font:
            args += ["--prop", f"headingFont={heading_font}"]
        if body_font:
            args += ["--prop", f"bodyFont={body_font}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_theme_fonts: 至少传一个字体")
        return self.runner.run(args)

    def set_presentation_props(self, *, title: str = "",
                               author: str = "", subject: str = "",
                               slide_size: str = "",
                               first_slide_num: int | None = None) -> str:
        """设置演示文稿级属性。

        参数:
            title/author/subject: 文档核心属性（文件信息里显示）。
            slide_size: 幻灯片尺寸预设 'widescreen'(16:9)/'standard'(4:3)/'a4' 等。
            first_slide_num: 起始页码（默认 1）。
        """
        args = ["set", self.doc_path, "/"]
        if title:
            args += ["--prop", f"title={title}"]
        if author:
            args += ["--prop", f"author={author}"]
        if subject:
            args += ["--prop", f"subject={subject}"]
        if slide_size:
            args += ["--prop", f"slideSize={slide_size}"]
        if first_slide_num is not None:
            args += ["--prop", f"firstSlideNum={first_slide_num}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_presentation_props: 至少传一个属性")
        return self.runner.run(args)

    # ---- 读取 ----
    def view_text(self) -> str:
        """读幻灯片文本（按 slide 分段，便于 LLM 自查）。"""
        return self.runner.run(["view", self.doc_path, "text"])

    def view_stats(self) -> Any:
        return self.runner.run(["view", self.doc_path, "stats"], json_output=True)

    def validate(self) -> str:
        return self.runner.run(["validate", self.doc_path])
