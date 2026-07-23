"""PowerPoint（.pptx）演示文稿操作工具：PptxTool。

绑定单个 .pptx 路径的结构化操作集合。幻灯片路径 ``/slide[N]``（1-based）。
shape 路径优先用 add 返回的 ``/slide[N]/shape[@id=ID]``（positional /shape[M]
会混入布局占位符）。表格单元格路径是 ``/slide[N]/table[M]/tr[R]/tc[C]``。
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from typing import Any

from .runner import OfficeCLIError, get_runner

logger = logging.getLogger(__name__)


@dataclass
class PptxTool:
    """绑定单个 .pptx 路径的 PowerPoint 操作工具。

    供 LangGraph 节点直接调用，内部翻译为 officecli 命令。
    """

    doc_path: str

    @property
    def runner(self) -> Any:
        return get_runner()

    # ---- 生命周期 ----
    def create(self) -> str:
        """创建空 pptx（若已存在会被覆盖）。"""
        return self.runner.run(["create", self.doc_path, "--force"])

    def close(self) -> str:
        return self.runner.run(["close", self.doc_path])

    # ---- 幻灯片 ----
    def add_slide(self, *, title: str = "", text: str = "", layout: str = "") -> str:
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
                if isinstance(data, dict)
                else []
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
    def add_textbox(
        self,
        slide_index: int,
        text: str,
        *,
        x: str = "1cm",
        y: str = "2cm",
        width: str = "22cm",
        height: str = "2cm",
        size: float | None = None,
        bold: bool = False,
        italic: bool = False,
        color: str = "",
        fill: str = "",
        align: str = "left",
        valign: str = "top",
    ) -> str:
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
            "add",
            self.doc_path,
            f"/slide[{slide_index}]",
            "--type",
            "textbox",
            "--prop",
            f"text={text}",
            "--prop",
            f"x={x}",
            "--prop",
            f"y={y}",
            "--prop",
            f"width={width}",
            "--prop",
            f"height={height}",
            "--prop",
            f"align={align}",
            "--prop",
            f"valign={valign}",
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

    def add_shape(
        self,
        slide_index: int,
        text: str,
        *,
        geometry: str = "rect",
        x: str = "2cm",
        y: str = "2cm",
        width: str = "4cm",
        height: str = "2cm",
        fill: str = "",
        line: str = "",
        size: float | None = None,
        bold: bool = False,
        color: str = "",
        align: str = "center",
        valign: str = "middle",
    ) -> str:
        """在指定幻灯片上加自选图形（矩形/椭圆/箭头等）+ 文字。

        参数:
            geometry: 形状预设，如 'rect'/'roundRect'/'ellipse'/'triangle'/
                      'diamond'/'rightArrow'/'star5'。
            line: 边框，格式 'color[:width[:style]]'（如 'FF0000:1.5:dash'）。
            其余参数同 add_textbox。
        """
        args = [
            "add",
            self.doc_path,
            f"/slide[{slide_index}]",
            "--type",
            "shape",
            "--prop",
            f"text={text}",
            "--prop",
            f"geometry={geometry}",
            "--prop",
            f"x={x}",
            "--prop",
            f"y={y}",
            "--prop",
            f"width={width}",
            "--prop",
            f"height={height}",
            "--prop",
            f"align={align}",
            "--prop",
            f"valign={valign}",
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
    def add_image(
        self,
        slide_index: int,
        src: str,
        *,
        x: str = "2cm",
        y: str = "2cm",
        width: str = "10cm",
        height: str = "",
        alt: str = "",
    ) -> str:
        """在指定幻灯片上插入图片。

        参数:
            slide_index: 幻灯片 1-based 序号。
            src: 图片来源（本地路径 / URL / data URI）。
            x/y/width/height: 位置和尺寸。height 留空则按图片比例。
            alt: 替代文本。
        """
        args = [
            "add",
            self.doc_path,
            f"/slide[{slide_index}]",
            "--type",
            "picture",
            "--prop",
            f"src={src}",
            "--prop",
            f"x={x}",
            "--prop",
            f"y={y}",
            "--prop",
            f"width={width}",
        ]
        if height:
            args += ["--prop", f"height={height}"]
        if alt:
            args += ["--prop", f"alt={alt}"]
        return self.runner.run(args)

    # ---- 表格 ----
    def add_table(
        self,
        slide_index: int,
        data: list[list],
        *,
        has_header: bool = True,
        x: str = "1cm",
        y: str = "4cm",
        width: str = "22cm",
    ) -> str:
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

        self.runner.run(
            [
                "add",
                self.doc_path,
                f"/slide[{slide_index}]",
                "--type",
                "table",
                "--prop",
                f"rows={rows}",
                "--prop",
                f"cols={cols}",
                "--prop",
                f"x={x}",
                "--prop",
                f"y={y}",
                "--prop",
                f"width={width}",
            ]
        )
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
                if isinstance(data, dict)
                else []
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
    def _build_table_ops(
        slide_index: int, tbl_index: int, norm: list[list[str]], has_header: bool
    ) -> list[dict]:
        """构造写 pptx 表格单元格的 batch ops。"""
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
                        "path": (
                            f"/slide[{slide_index}]/table[{tbl_index}]/tr[{r + 1}]/tc[{c + 1}]"
                        ),
                        "props": props,
                    }
                )
        return ops

    # ---- 通用 ----
    def batch(self, ops: list[dict]) -> Any:
        """批量原子操作。复用 DocTool 同款实现（argv 传 JSON）。"""
        payload = json.dumps(ops, ensure_ascii=False)
        return self.runner.run(
            ["batch", self.doc_path, "--commands", payload, "--json"],
        )

    # ---- 幻灯片级元数据（与占位符模式兼容，不会重叠）----
    def set_transition(
        self,
        slide_index: int,
        transition: str,
        *,
        advance_time_ms: int | None = None,
        advance_on_click: bool = True,
    ) -> str:
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
        args = [
            "set",
            self.doc_path,
            f"/slide[{slide_index}]",
            "--prop",
            f"transition={transition}",
        ]
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
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/slide[{slide_index}]",
                "--prop",
                f"notes={text}",
            ]
        )

    def set_slide_hidden(self, slide_index: int, hidden: bool = True) -> str:
        """隐藏/显示幻灯片（放映时跳过）。"""
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/slide[{slide_index}]",
                "--prop",
                f"hidden={'true' if hidden else 'false'}",
            ]
        )

    # ---- 演示文稿级（deck-wide）----
    def set_theme_colors(
        self,
        *,
        accent1: str = "",
        accent2: str = "",
        accent3: str = "",
        accent4: str = "",
        accent5: str = "",
        accent6: str = "",
        hyperlink: str = "",
    ) -> str:
        """自定义主题颜色（影响全 deck 的强调色/超链接色）。

        参数: 6 个 accent 颜色 + hyperlink，6位十六进制（如 '4472C4'）。
        只传需要改的，其余保持不变。
        """
        args = ["set", self.doc_path, "/theme"]
        for k, v in [
            ("accent1", accent1),
            ("accent2", accent2),
            ("accent3", accent3),
            ("accent4", accent4),
            ("accent5", accent5),
            ("accent6", accent6),
            ("hyperlink", hyperlink),
        ]:
            if v:
                args += ["--prop", f"{k}={v}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_theme_colors: 至少传一个颜色")
        return self.runner.run(args)

    def set_theme_fonts(self, heading_font: str = "", body_font: str = "") -> str:
        """自定义主题字体（标题字体 / 正文字体，影响全 deck）。"""
        args = ["set", self.doc_path, "/theme"]
        if heading_font:
            args += ["--prop", f"headingFont={heading_font}"]
        if body_font:
            args += ["--prop", f"bodyFont={body_font}"]
        if len(args) <= 3:
            raise OfficeCLIError("set_theme_fonts: 至少传一个字体")
        return self.runner.run(args)

    def set_presentation_props(
        self,
        *,
        title: str = "",
        author: str = "",
        subject: str = "",
        slide_size: str = "",
        first_slide_num: int | None = None,
    ) -> str:
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
