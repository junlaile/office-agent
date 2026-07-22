"""Word（.docx）文档操作工具：DocTool。

绑定单个 .docx 路径的结构化操作集合，供 LangGraph 节点直接调用。
所有格式用显式 props（size/bold/color/listStyle）确保 Word 真实生效，
不依赖默认 docx 缺失的命名样式（Heading1/Title 等不存在）。

文本一律通过 argv 参数数组传递，绝不拼 shell 字符串。
写入优先用 batch（--commands JSON），失败整体回滚；add_table 内部用 batch 写单元格。
表格元素路径是 /body/tbl[N]/tr[N]/tc[N]（注意是 tbl/tr/tc，不是 table/row/cell）。
"""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .cli_runner import OfficeCLIError, get_runner

logger = logging.getLogger(__name__)


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
    def runner(self) -> Any:
        return get_runner()

    # ---- 生命周期 ----
    def create(self) -> str:
        """创建空文档。若已存在（如公文模式已复制模板）则跳过，不覆盖。

        公文模式下，main.py 或 start_from_template 会先把模板复制到
        doc_path；此时 LLM 若仍调 create_doc，应保留模板而非覆盖成空白。
        """
        if Path(self.doc_path).exists():
            return f"文档已存在，跳过创建（保留已有内容）: {self.doc_path}"
        return self.runner.run(["create", self.doc_path, "--force"])

    def close(self) -> str:
        """刷盘并释放（关闭常驻会话）。"""
        return self.runner.run(["close", self.doc_path])

    # ---- 写入 ----
    def add_title(self, text: str) -> str:
        """加文档主标题。显式大字号 + 加粗，居中。"""
        return self.runner.run(
            [
                "add",
                self.doc_path,
                "/body",
                "--type",
                "paragraph",
                "--prop",
                f"text={text}",
                "--prop",
                f"size={_TITLE_SIZE_PT}",
                "--prop",
                "bold=true",
                "--prop",
                "align=center",
            ]
        )

    def add_heading(self, text: str, level: int = 1) -> str:
        """加章节标题。level 1-9，字号递减、加粗。"""
        level = max(1, min(9, int(level)))
        size = _HEADING_SIZE_PT[level]
        return self.runner.run(
            [
                "add",
                self.doc_path,
                "/body",
                "--type",
                "paragraph",
                "--prop",
                f"text={text}",
                "--prop",
                f"size={size}",
                "--prop",
                "bold=true",
                "--prop",
                "spaceBefore=12pt",
                "--prop",
                "spaceAfter=6pt",
            ]
        )

    def add_paragraph(
        self, text: str, *, bold: bool = False, italic: bool = False, size: float | None = None
    ) -> str:
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
        return self.runner.run(
            [
                "add",
                self.doc_path,
                "/body",
                "--type",
                "paragraph",
                "--prop",
                f"text={text}",
                "--prop",
                f"listStyle={style}",
            ]
        )

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
        add_output = self.runner.run(
            [
                "add",
                self.doc_path,
                "/body",
                "--type",
                "table",
                "--prop",
                f"rows={rows}",
                "--prop",
                f"cols={cols}",
            ]
        )
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

        m = re.search(r"tbl\[(\d+)\]", output or "")
        return int(m.group(1)) if m else 0

    @staticmethod
    def _build_table_ops(norm: list[list[str]], tbl_index: int, has_header: bool) -> list[dict]:
        """构造写表格单元格的 batch ops。tbl_index 1-based。"""
        if tbl_index <= 0:
            tbl_index = 1
        ops: list[dict] = []
        for r, row in enumerate(norm):
            for c, cell in enumerate(row):
                props: dict[str, str] = {"text": str(cell)}
                if has_header and r == 0:
                    props["bold"] = "true"
                ops.append(
                    {
                        "command": "set",
                        "path": f"/body/tbl[{tbl_index}]/tr[{r + 1}]/tc[{c + 1}]",
                        "props": props,
                    }
                )
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
                if isinstance(data, dict)
                else []
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

    def add_image(self, src: str, *, width: str = "8cm", alt: str = "", caption: str = "") -> str:
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
        self.runner.run(
            [
                "add",
                self.doc_path,
                "/body",
                "--type",
                "paragraph",
            ]
        )
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
                if isinstance(data, dict)
                else []
            )
            p_indices = [
                int(p.split("p[")[1].rstrip("]"))
                for c in children
                if (p := c.get("path", "")) and "/p[" in p
            ]
            return max(p_indices) if p_indices else 0
        except Exception:  # noqa: BLE001
            return 0

    # ---- 编辑（改/删/替换）----
    def set_paragraph_text(self, path: str, text: str) -> str:
        """整段重写文字。

        用 ``set <doc> <path> --prop text=<text>`` 实现，会清空原段落所有
        run、新建单个 implicit run。后果:

        - 【丢失】原段内逐 run 的字体/字号/粗体/颜色（run 级 rPr）。
        - 【保留】段落级属性（对齐/缩进/行距/段前后/style 等，pPr 上的）。

        所以对公文标题段（小标宋 22pt）、正文段（仿宋 16pt），用它会把
        字体重置成默认——只适合"整段全换且不介意字体"的场景。
        想改字的同时保字体，用 :meth:`find_replace`。

        参数:
            path: 段落路径，如 ``/body/p[4]``（来自 view_text 的标注）。
            text: 新的整段文字（纯文本）。
        """
        if not path or not text:
            raise OfficeCLIError("set_paragraph_text: path 和 text 不能为空")
        return self.runner.run(
            [
                "set",
                self.doc_path,
                path,
                "--prop",
                f"text={text}",
            ]
        )

    def find_replace(
        self, find: str, replace: str, *, path: str = "/body", regex: bool = False
    ) -> str:
        """子串替换（保留段落内字体格式，首选编辑方式）。

        用 ``set <doc> <path> --find <find> --replace <replace>`` 实现。
        只替换匹配的子串，段落其余文字和 run 级格式（字体/字号/粗体）保留。
        适合把模板正文里的 'XX'、'XX工作' 等占位换成真实内容。

        参数:
            find: 要查找的文字。默认字面子串匹配。
            replace: 替换成的文字。
            path: 作用域。``/body`` = 全文 body（默认）；
                  ``/body/p[N]`` = 仅该段。
            regex: True 时把 find 当正则（officecli 用 ``r"..."`` 前缀启用）。
                   例 find=r"\\d{4}年" 匹配"2026年"。

        返回 officecli 的 stdout（含替换计数）。
        """
        if not find:
            raise OfficeCLIError("find_replace: find 不能为空")
        pattern = f'r"{find}"' if regex else find
        return self.runner.run(
            [
                "set",
                self.doc_path,
                path,
                "--find",
                pattern,
                "--replace",
                replace,
            ]
        )

    def remove(self, path: str) -> str:
        """删除指定路径的元素（段落/表格/图片等）。

        用 ``remove <doc> <path>``。删段后，后续段落索引立即前移
        （/body/p[3] 删掉后，原 p[4] 变成新的 p[3]）。连续删多段时，
        建议从后往前删，或每删一段重新 view_text 确认新索引。

        参数:
            path: 要删除的元素路径，如 ``/body/p[10]``、``/body/tbl[1]``。
        """
        if not path:
            raise OfficeCLIError("remove: path 不能为空")
        return self.runner.run(["remove", self.doc_path, path])

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
    def add_toc(
        self,
        *,
        levels: str = "1-3",
        title: str = "",
        hyperlinks: bool = True,
        page_numbers: bool = True,
    ) -> str:
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
            "add",
            self.doc_path,
            "/",
            "--type",
            "toc",
            "--prop",
            f"levels={levels}",
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
    def add_header(
        self, text: str = "", *, field: str = "", align: str = "", kind: str = "default"
    ) -> str:
        """加页眉。

        参数:
            text: 页眉文字（与 field 二选一或组合）。
            field: 域类型，如 'page'(页码)/'numpages'(总页数)/'date'/'author'/'title'。
            align: 对齐 left/center/right。
            kind: 'default'(所有页)/'first'(首页不同)/'even'(偶数页不同)。
        """
        args = ["add", self.doc_path, "/", "--type", "header", "--prop", f"type={kind}"]
        if text:
            args += ["--prop", f"text={text}"]
        if field:
            args += ["--prop", f"field={field}"]
        if align:
            args += ["--prop", f"align={align}"]
        return self.runner.run(args)

    def add_footer(
        self, text: str = "", *, field: str = "", align: str = "center", kind: str = "default"
    ) -> str:
        """加页脚。参数同 add_header，默认居中。"""
        args = ["add", self.doc_path, "/", "--type", "footer", "--prop", f"type={kind}"]
        if text:
            args += ["--prop", f"text={text}"]
        if field:
            args += ["--prop", f"field={field}"]
        if align:
            args += ["--prop", f"align={align}"]
        return self.runner.run(args)

    # ---- 超链接 ----
    def add_hyperlink(
        self, text: str, url: str = "", *, anchor: str = "", tooltip: str = ""
    ) -> str:
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
        args = [
            "add",
            self.doc_path,
            f"/body/p[{p_index}]",
            "--type",
            "hyperlink",
            "--prop",
            f"text={text}",
        ]
        if url:
            args += ["--prop", f"url={url}"]
        elif anchor:
            args += ["--prop", f"anchor={anchor}"]
        if tooltip:
            args += ["--prop", f"tooltip={tooltip}"]
        return self.runner.run(args)

    # ---- 图表 ----
    def add_chart(
        self,
        chart_type: str,
        data: str,
        *,
        categories: str = "",
        title: str = "",
        width: str = "15cm",
        height: str = "8cm",
        legend: str = "bottom",
    ) -> str:
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
            "add",
            self.doc_path,
            f"/body/p[{p_index}]",
            "--type",
            "chart",
            "--prop",
            f"chartType={chart_type}",
            "--prop",
            f"data={data}",
            "--prop",
            f"width={width}",
            "--prop",
            f"height={height}",
            "--prop",
            f"legend={legend}",
        ]
        if categories:
            args += ["--prop", f"categories={categories}"]
        if title:
            args += ["--prop", f"title={title}"]
        return self.runner.run(args)

    # ---- 分节 ----
    def add_section(
        self, *, section_type: str = "nextPage", orientation: str = "", columns: int = 1
    ) -> str:
        """插入分节符。

        参数:
            section_type: 'nextPage'(下一页)/'continuous'(连续)/'evenPage'/'oddPage'。
            orientation: 'portrait'/'landscape'（留空不变）。
            columns: 分栏数（1=不分栏）。
        """
        args = ["add", self.doc_path, "/", "--type", "section", "--prop", f"type={section_type}"]
        if orientation:
            args += ["--prop", f"orientation={orientation}"]
        if columns > 1:
            args += ["--prop", f"columns={columns}"]
        return self.runner.run(args)

    # ---- 文档属性 ----
    def set_doc_properties(
        self,
        *,
        title: str = "",
        author: str = "",
        subject: str = "",
        keywords: str = "",
        description: str = "",
    ) -> str:
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
        return self.runner.run(
            [
                "add",
                self.doc_path,
                "/body",
                "--type",
                "paragraph",
                "--prop",
                f"text={text}",
                "--prop",
                f"style={style}",
            ]
        )

    def add_style(
        self,
        style_id: str,
        name: str,
        *,
        style_type: str = "paragraph",
        based_on: str = "Normal",
        size: float | None = None,
        bold: bool = False,
        color: str = "",
        outline_level: int | None = None,
    ) -> str:
        """创建命名样式。

        参数:
            style_id: 样式 ID（如 'MyHeading'）。
            name: 显示名。
            style_type: 'paragraph'/'character'/'table'。
            based_on: 继承自哪个样式（默认 Normal）。
            outline_level: 大纲级别 0-9（0=一级标题，能让 TOC 收录）。
        """
        args = [
            "add",
            self.doc_path,
            "/styles",
            "--type",
            "style",
            "--prop",
            f"id={style_id}",
            "--prop",
            f"name={name}",
            "--prop",
            f"type={style_type}",
            "--prop",
            f"basedOn={based_on}",
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
