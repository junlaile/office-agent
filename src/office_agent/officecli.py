"""OfficeCLI 的 Python 封装：subprocess 调用 + 结构化工具集。

经实测验证的关键点（officecli v1.0.139）:
    - subprocess 必须用 encoding='utf-8'，且 JSON 通过 --commands argv 传递
      （stdin 传中文 JSON 会被误解析；shell echo 也会转码）。
    - 默认 create 的 docx 只有 Normal 样式，Heading1/Title/ListBullet 不存在，
      因此本封装层用【显式格式 props】(size/bold/color/font/listStyle/outlineLvl)
      而非命名样式，确保 Word 打开时格式真实生效。
    - 表格元素路径是 /body/tbl[N]/tr[N]/tc[N]（注意是 tbl/tr/tc，不是 table/row/cell）。
    - 设置 OFFICECLI_NO_AUTO_RESIDENT=1 避免后台常驻进程文件锁。
    - 图片 width/height 必须带单位（cm/in/pt/px/mm），裸数字会被当成 EMU。

设计:
    - 文本一律通过 argv 参数数组传递，绝不拼 shell 字符串，杜绝注入。
    - 写入优先用 batch（--commands JSON），失败整体回滚；add_table 内部用 batch 写单元格。
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
from dataclasses import dataclass, field
from typing import Any

from .config import settings

# 图片尺寸必须带单位，避免裸数字被当成 EMU
_SIZE_UNIT_RE = re.compile(r"^\d+(\.\d+)?(cm|in|pt|px|mm)$", re.IGNORECASE)
_ALIGN_OK = {"left", "center", "right", "justify", "both", "distribute"}


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


def _validate_size_unit(value: str, *, label: str = "尺寸") -> str:
    """校验带单位的尺寸字符串；非法则抛 OfficeCLIError。"""
    v = (value or "").strip()
    if not v:
        raise OfficeCLIError(f"{label}不能为空，需带单位如 8cm / 400px / 3in")
    if not _SIZE_UNIT_RE.match(v):
        raise OfficeCLIError(
            f"{label}无效: {value!r}。必须带单位，例如 8cm / 400px / 3in / 12pt "
            f"（裸数字会被当成 EMU，导致图片异常巨大）",
        )
    return v


def _normalize_align(align: str | None) -> str | None:
    if not align:
        return None
    a = align.strip().lower()
    if a not in _ALIGN_OK:
        raise OfficeCLIError(
            f"align 无效: {align!r}。允许: {', '.join(sorted(_ALIGN_OK))}",
        )
    return a


@dataclass
class DocTool:
    """绑定单个文档路径的结构化操作工具。

    供 LangGraph 节点直接调用，语义清晰，内部翻译为 officecli 命令。
    所有格式用显式 props（size/bold/color/listStyle/outlineLvl）确保 Word 真实生效，
    不依赖默认 docx 缺失的命名样式。
    """

    doc_path: str

    @property
    def runner(self) -> _Runner:
        return get_runner()

    # ---- 生命周期 ----
    def create(self, *, locale: str = "zh-CN") -> str:
        """创建空文档（若已存在会被覆盖）。默认 zh-CN 以稳住中文字体。"""
        args = ["create", self.doc_path, "--force"]
        if locale:
            args += ["--locale", locale]
        return self.runner.run(args)

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
        """加章节标题。level 1-9，字号递减、加粗，并写 outlineLvl 便于大纲视图。"""
        level = max(1, min(9, int(level)))
        size = _HEADING_SIZE_PT[level]
        # Word outlineLvl 0-based：一级标题 = 0
        outline = level - 1
        return self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", f"text={text}",
            "--prop", f"size={size}",
            "--prop", "bold=true",
            "--prop", f"outlineLvl={outline}",
            "--prop", "spaceBefore=12pt",
            "--prop", "spaceAfter=6pt",
        ])

    def add_paragraph(
        self,
        text: str,
        *,
        bold: bool = False,
        italic: bool = False,
        size: float | None = None,
        align: str | None = None,
    ) -> str:
        """加正文段落。"""
        props = [f"text={text}"]
        if bold:
            props.append("bold=true")
        if italic:
            props.append("italic=true")
        if size is not None:
            props.append(f"size={size}")
        a = _normalize_align(align)
        if a:
            props.append(f"align={a}")
        args = ["add", self.doc_path, "/body", "--type", "paragraph"]
        for p in props:
            args += ["--prop", p]
        return self.runner.run(args)

    def add_list_item(
        self,
        text: str,
        *,
        ordered: bool = False,
        level: int = 0,
    ) -> str:
        """加列表项。用 listStyle=bullet/ordered；level 为嵌套层级 0-8。"""
        style = "ordered" if ordered else "bullet"
        lvl = max(0, min(8, int(level)))
        args = [
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", f"text={text}",
            "--prop", f"listStyle={style}",
        ]
        if lvl > 0:
            args += ["--prop", f"numLevel={lvl}"]
        return self.runner.run(args)

    def add_table(
        self,
        data: list[list[str]],
        *,
        has_header: bool = True,
        style: str | None = "medium2",
    ) -> Any:
        """加表格。data 是二维字符串数组。

        实现: add table 建空表 → 查询新表真实索引 → batch 逐单元格写入。

        为什么不用 --prop data 一步创建: officecli 的 data CSV 格式【不支持
        引号转义】，单元格内含逗号/分号会被拆成多列（如"50,000"→两列），
        导致列数错乱。逐单元格 batch 写入对任意内容都安全。
        """
        if not data or not data[0]:
            raise OfficeCLIError("表格数据为空")

        rows = len(data)
        cols = max(len(row) for row in data)
        norm = [list(row) + [""] * (cols - len(row)) for row in data]

        add_args = [
            "add", self.doc_path, "/body", "--type", "table",
            "--prop", f"rows={rows}",
            "--prop", f"cols={cols}",
        ]
        if style:
            add_args += ["--prop", f"style={style}"]
        add_output = self.runner.run(add_args)
        tbl_index = self._parse_tbl_index(add_output) or self._last_table_index()

        ops = self._build_table_ops(norm, tbl_index, has_header)
        try:
            self.batch(ops)
        except OfficeCLIError:
            tbl_index = self._last_table_index()
            ops2 = self._build_table_ops(norm, tbl_index, has_header)
            try:
                self.batch(ops2)
            except OfficeCLIError:
                failed: list[str] = []
                for op in ops2:
                    try:
                        self.batch([op])
                    except OfficeCLIError as e:
                        failed.append(f"{op.get('path')}: {e}")
                if failed:
                    preview = "; ".join(failed[:5])
                    more = f" 等共 {len(failed)} 处" if len(failed) > 5 else ""
                    raise OfficeCLIError(
                        f"表格部分单元格写入失败（{len(failed)}/{len(ops2)}）: "
                        f"{preview}{more}",
                    ) from None

        return f"已添加 {rows} 行 × {cols} 列的表格"

    @staticmethod
    def _parse_tbl_index(output: str) -> int:
        """从 'add table' 的输出 'Added table at /body/tbl[N]' 解析索引 N。"""
        m = re.search(r"tbl\[(\d+)\]", output or "")
        return int(m.group(1)) if m else 0

    @staticmethod
    def _parse_path_index(output: str, tag: str) -> int:
        """从 add 输出解析 /body/{tag}[N] 索引。"""
        m = re.search(rf"{re.escape(tag)}\[(\d+)\]", output or "")
        return int(m.group(1)) if m else 0

    @staticmethod
    def _build_table_ops(
        norm: list[list[str]],
        tbl_index: int,
        has_header: bool,
    ) -> list[dict]:
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
        """批量原子操作。ops 见 officecli batch 文档。"""
        payload = json.dumps(ops, ensure_ascii=False)
        return self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )

    def add_image(
        self,
        src: str,
        *,
        width: str = "8cm",
        height: str = "",
        alt: str = "",
        caption: str = "",
    ) -> str:
        """插入图片到文档末尾。picture 的 parent 必须是 paragraph。"""
        width = _validate_size_unit(width, label="width")
        if height:
            height = _validate_size_unit(height, label="height")

        add_p_out = self.runner.run([
            "add", self.doc_path, "/body", "--type", "paragraph",
            "--prop", "align=center",
        ])
        p_index = (
            self._parse_path_index(add_p_out, "p")
            or self._last_paragraph_index()
            or 1
        )

        props = [f"src={src}", f"width={width}"]
        if height:
            props.append(f"height={height}")
        if alt:
            props.append(f"alt={alt}")
        args = ["add", self.doc_path, f"/body/p[{p_index}]", "--type", "picture"]
        for p in props:
            args += ["--prop", p]
        self.runner.run(args)

        if caption:
            self.add_paragraph(caption, italic=True, align="center")

        extra = f"×{height}" if height else ""
        return f"已插入图片（{width}{extra}）" + (
            f"，图注: {caption}" if caption else ""
        )

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

    # ---- 页眉 / 页脚 / 分页 / 替换 ----
    def add_header(
        self,
        text: str = "",
        *,
        align: str = "center",
        page_numbers: bool = False,
    ) -> str:
        """添加页眉。page_numbers=True 时组装「第 X 页 / 共 Y 页」。"""
        return self._add_header_or_footer(
            "header", text=text, align=align, page_numbers=page_numbers,
        )

    def add_footer(
        self,
        text: str = "",
        *,
        align: str = "center",
        page_numbers: bool = False,
    ) -> str:
        """添加页脚。page_numbers=True 时组装「第 X 页 / 共 Y 页」。"""
        return self._add_header_or_footer(
            "footer", text=text, align=align, page_numbers=page_numbers,
        )

    def _add_header_or_footer(
        self,
        kind: str,
        *,
        text: str,
        align: str,
        page_numbers: bool,
    ) -> str:
        a = _normalize_align(align) or "center"
        if page_numbers:
            # 按 officecli 文档分步组装 Page X of Y
            prefix = (text.strip() + " " if text.strip() else "") + "第 "
            add_out = self.runner.run([
                "add", self.doc_path, "/", "--type", kind,
                "--prop", f"text={prefix}",
                "--prop", f"align={a}",
            ])
            idx = self._parse_path_index(add_out, kind) or 1
            parent = f"/{kind}[{idx}]/p[1]"
            self.runner.run([
                "add", self.doc_path, parent, "--type", "field",
                "--prop", "fieldType=page",
            ])
            self.runner.run([
                "add", self.doc_path, parent, "--type", "run",
                "--prop", "text= 页 / 共 ",
            ])
            self.runner.run([
                "add", self.doc_path, parent, "--type", "field",
                "--prop", "fieldType=numpages",
            ])
            self.runner.run([
                "add", self.doc_path, parent, "--type", "run",
                "--prop", "text= 页",
            ])
            return f"已添加{('页眉' if kind == 'header' else '页脚')}（含页码）"

        if not text.strip():
            raise OfficeCLIError(
                f"{'页眉' if kind == 'header' else '页脚'}文字为空，"
                f"且未开启 page_numbers",
            )
        self.runner.run([
            "add", self.doc_path, "/", "--type", kind,
            "--prop", f"text={text}",
            "--prop", f"align={a}",
        ])
        return f"已添加{('页眉' if kind == 'header' else '页脚')}: {text}"

    def add_page_break(self) -> str:
        """在正文末尾插入分页符。"""
        self.runner.run([
            "add", self.doc_path, "/body", "--type", "pagebreak",
        ])
        return "已插入分页符"

    def replace_text(self, find: str, replace: str) -> str:
        """在正文中查找替换（字面量子串）。"""
        if not find:
            raise OfficeCLIError("查找文本 find 不能为空")
        out = self.runner.run([
            "set", self.doc_path, "/body",
            "--find", find,
            "--replace", replace if replace is not None else "",
        ])
        return (out or "").strip() or f"已替换: {find!r} → {replace!r}"

    def batch_add(self, ops: list[dict]) -> str:
        """把高层 ops 编译为一次 officecli batch 提交。

        支持的 op:
          - title: {op, text}
          - heading: {op, text, level?}
          - paragraph: {op, text, bold?, italic?, size?, align?}
          - list_item: {op, text, ordered?, level?}
          - page_break: {op}
        表格/图片需走专用方法（依赖索引回读），不可放进 batch_add。
        """
        if not ops:
            raise OfficeCLIError("batch_add 操作列表为空")

        commands: list[dict] = []
        for i, raw in enumerate(ops):
            if not isinstance(raw, dict):
                raise OfficeCLIError(f"batch_add ops[{i}] 必须是对象")
            op = (raw.get("op") or raw.get("type") or "").strip().lower()
            if not op:
                raise OfficeCLIError(f"batch_add ops[{i}] 缺少 op 字段")
            commands.append(self._compile_batch_op(op, raw, i))

        self.batch(commands)
        return f"已批量写入 {len(commands)} 项"

    def _compile_batch_op(self, op: str, raw: dict, index: int) -> dict:
        if op == "page_break":
            return {"command": "add", "parent": "/body", "type": "pagebreak"}

        text = raw.get("text")
        if text is None or str(text) == "":
            raise OfficeCLIError(f"batch_add ops[{index}] ({op}) 缺少 text")

        props: dict[str, str] = {"text": str(text)}

        if op == "title":
            props.update({
                "size": str(_TITLE_SIZE_PT),
                "bold": "true",
                "align": "center",
            })
        elif op == "heading":
            level = max(1, min(9, int(raw.get("level", 1))))
            props.update({
                "size": str(_HEADING_SIZE_PT[level]),
                "bold": "true",
                "outlineLvl": str(level - 1),
                "spaceBefore": "12pt",
                "spaceAfter": "6pt",
            })
        elif op == "paragraph":
            if raw.get("bold"):
                props["bold"] = "true"
            if raw.get("italic"):
                props["italic"] = "true"
            if raw.get("size") is not None:
                props["size"] = str(raw["size"])
            a = _normalize_align(raw.get("align") or None)
            if a:
                props["align"] = a
        elif op == "list_item":
            ordered = bool(raw.get("ordered", False))
            props["listStyle"] = "ordered" if ordered else "bullet"
            lvl = max(0, min(8, int(raw.get("level", 0))))
            if lvl > 0:
                props["numLevel"] = str(lvl)
        else:
            raise OfficeCLIError(
                f"batch_add ops[{index}] 不支持 op={op!r}。"
                f"允许: title/heading/paragraph/list_item/page_break",
            )

        return {
            "command": "add",
            "parent": "/body",
            "type": "paragraph",
            "props": props,
        }

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


# ============================================================
# Excel: SheetTool
# ============================================================
def _col_letters(index: int) -> str:
    """0-based 列索引 → A, B, ..., Z, AA, ..."""
    n = index + 1
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _parse_a1(ref: str) -> tuple[int, int]:
    """解析 A1 样式引用 → (row0, col0)。"""
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", (ref or "").strip())
    if not m:
        raise OfficeCLIError(f"无效单元格引用: {ref!r}（期望如 A1）")
    col_s, row_s = m.group(1).upper(), m.group(2)
    col = 0
    for ch in col_s:
        col = col * 26 + (ord(ch) - 64)
    return int(row_s) - 1, col - 1


@dataclass
class SheetTool:
    """绑定单个 .xlsx 路径的 Excel MVP 工具。"""

    doc_path: str

    @property
    def runner(self) -> _Runner:
        return get_runner()

    def create(self) -> str:
        return self.runner.run(["create", self.doc_path, "--force"])

    def add_sheet(self, name: str) -> str:
        name = (name or "").strip()
        if not name:
            raise OfficeCLIError("工作表名称不能为空")
        return self.runner.run([
            "add", self.doc_path, "/", "--type", "sheet",
            "--prop", f"name={name}",
        ])

    def write_cell(
        self,
        sheet: str,
        ref: str,
        value: str = "",
        *,
        formula: str = "",
        bold: bool = False,
    ) -> str:
        sheet = (sheet or "Sheet1").strip() or "Sheet1"
        ref = (ref or "").strip()
        if not ref:
            raise OfficeCLIError("单元格引用 ref 不能为空")
        _parse_a1(ref)  # validate
        path = f"/{sheet}/{ref}"
        props: list[str] = []
        if formula:
            # officecli: formula 不带前导 =
            f = formula[1:] if formula.startswith("=") else formula
            props.append(f"formula={f}")
        else:
            props.append(f"value={value}")
        if bold:
            props.append("bold=true")
        args = ["set", self.doc_path, path]
        for p in props:
            args += ["--prop", p]
        out = self.runner.run(args)
        return (out or "").strip() or f"已写入 {path}"

    def write_range(
        self,
        sheet: str,
        start: str,
        data: list[list[str]],
        *,
        header_bold: bool = True,
    ) -> str:
        """从 start（如 A1）起写入二维数据。"""
        if not data:
            raise OfficeCLIError("写入数据为空")
        sheet = (sheet or "Sheet1").strip() or "Sheet1"
        row0, col0 = _parse_a1(start or "A1")
        ops: list[dict] = []
        for r, row in enumerate(data):
            if row is None:
                continue
            for c, cell in enumerate(row):
                ref = f"{_col_letters(col0 + c)}{row0 + r + 1}"
                props: dict[str, str] = {"value": "" if cell is None else str(cell)}
                if header_bold and r == 0:
                    props["bold"] = "true"
                ops.append({
                    "command": "set",
                    "path": f"/{sheet}/{ref}",
                    "props": props,
                })
        if not ops:
            raise OfficeCLIError("写入数据为空")
        self.batch(ops)
        rows = len(data)
        cols = max(len(r) for r in data if r is not None)
        return f"已写入 {sheet}!{start} 起 {rows} 行 × {cols} 列"

    def batch(self, ops: list[dict]) -> Any:
        payload = json.dumps(ops, ensure_ascii=False)
        return self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )

    def view_text(self) -> str:
        return self.runner.run(["view", self.doc_path, "text"])

    def validate(self) -> str:
        return self.runner.run(["validate", self.doc_path])


# ============================================================
# PowerPoint: SlideTool
# ============================================================
@dataclass
class SlideTool:
    """绑定单个 .pptx 路径的 PowerPoint MVP 工具。"""

    doc_path: str

    @property
    def runner(self) -> _Runner:
        return get_runner()

    def create(self) -> str:
        return self.runner.run(["create", self.doc_path, "--force"])

    def add_slide(
        self,
        title: str = "",
        body: str = "",
        *,
        layout: str = "",
    ) -> str:
        args = ["add", self.doc_path, "/", "--type", "slide"]
        if layout:
            args += ["--prop", f"layout={layout}"]
        if title:
            args += ["--prop", f"title={title}"]
        if body:
            args += ["--prop", f"text={body}"]
        out = self.runner.run(args)
        idx = DocTool._parse_path_index(out, "slide") or 0
        return (out or "").strip() or (
            f"已添加幻灯片{f' /slide[{idx}]' if idx else ''}"
        )

    def add_bullets(self, slide_index: int, items: list[str]) -> str:
        """在指定页追加一个要点列表形状。"""
        clean = [str(i).strip() for i in (items or []) if str(i).strip()]
        if not clean:
            raise OfficeCLIError("要点列表为空")
        idx = max(1, int(slide_index))
        parent = f"/slide[{idx}]"
        add_out = self.runner.run([
            "add", self.doc_path, parent, "--type", "shape",
            "--prop", "geometry=rect",
            "--prop", "x=1.5cm",
            "--prop", "y=5cm",
            "--prop", "width=22cm",
            "--prop", "height=10cm",
        ])
        # 输出形如 shape[@id=100000]
        m = re.search(r"shape\[@id=(\d+)\]", add_out or "")
        shape_id = int(m.group(1)) if m else 0
        if shape_id <= 0:
            raise OfficeCLIError(f"无法定位新建形状: {add_out}")
        shape_path = f"{parent}/shape[@id={shape_id}]"
        for text in clean:
            self.runner.run([
                "add", self.doc_path, shape_path, "--type", "paragraph",
                "--prop", f"text={text}",
                "--prop", "list=bullet",
            ])
        return f"已在幻灯片 {idx} 添加 {len(clean)} 条要点"

    def add_table(self, slide_index: int, data: list[list[str]]) -> str:
        if not data or not data[0]:
            raise OfficeCLIError("表格数据为空")
        idx = max(1, int(slide_index))
        rows = len(data)
        cols = max(len(r) for r in data)
        norm = [
            [("" if c is None else str(c)) for c in (list(row) + [""] * (cols - len(row)))]
            for row in data
        ]
        add_out = self.runner.run([
            "add", self.doc_path, f"/slide[{idx}]", "--type", "table",
            "--prop", f"rows={rows}",
            "--prop", f"cols={cols}",
            "--prop", "x=1.5cm",
            "--prop", "y=5cm",
            "--prop", "width=22cm",
            "--prop", f"height={max(2, rows) * 1.2}cm",
        ])
        m = re.search(r"table\[@id=(\d+)\]", add_out or "")
        if not m:
            raise OfficeCLIError(f"无法定位新建表格: {add_out}")
        tid = m.group(1)
        ops = []
        for r, row in enumerate(norm):
            for c, cell in enumerate(row):
                ops.append({
                    "command": "set",
                    "path": (
                        f"/slide[{idx}]/table[@id={tid}]"
                        f"/tr[{r + 1}]/tc[{c + 1}]"
                    ),
                    "props": {"text": cell},
                })
        payload = json.dumps(ops, ensure_ascii=False)
        self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )
        return f"已在幻灯片 {idx} 添加 {rows}×{cols} 表格"

    def add_image(
        self,
        slide_index: int,
        src: str,
        *,
        width: str = "12cm",
    ) -> str:
        width = _validate_size_unit(width, label="width")
        idx = max(1, int(slide_index))
        self.runner.run([
            "add", self.doc_path, f"/slide[{idx}]", "--type", "picture",
            "--prop", f"src={src}",
            "--prop", f"width={width}",
            "--prop", "x=3cm",
            "--prop", "y=5cm",
        ])
        return f"已在幻灯片 {idx} 插入图片（{width}）"

    def view_outline(self) -> str:
        return self.runner.run(["view", self.doc_path, "outline"])

    def view_text(self) -> str:
        return self.runner.run(["view", self.doc_path, "text"])

    def validate(self) -> str:
        return self.runner.run(["validate", self.doc_path])
