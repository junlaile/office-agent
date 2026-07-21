"""Excel（.xlsx）工作簿操作工具：ExcelTool。

绑定单个 .xlsx 路径的结构化操作集合。单元格路径 ``/<sheetName>/<A1Ref>``。
写单元格一律用 ``set``（officecli 的 set 会自动创建不存在的单元格，add 不会）。
批量写用 batch + 逐 cell 的 set op，避免 officecli CSV data 的引号转义陷阱。
"""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from typing import Any

from .cli_runner import OfficeCLIError, get_runner

logger = logging.getLogger(__name__)


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
    def runner(self) -> Any:
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
        args = ["add", self.doc_path, "/", "--type", "sheet", "--prop", f"name={name}"]
        if tab_color:
            args += ["--prop", f"tabColor={tab_color}"]
        return self.runner.run(args)

    def rename_sheet(self, old: str, new: str) -> str:
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{old}",
                "--prop",
                f"name={new}",
            ]
        )

    def set_sheet_color(self, name: str, color: str) -> str:
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{name}",
                "--prop",
                f"tabColor={color}",
            ]
        )

    def set_doc_properties(
        self,
        *,
        title: str = "",
        author: str = "",
        subject: str = "",
        keywords: str = "",
        description: str = "",
    ) -> str:
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
    def set_cell(
        self,
        sheet: str,
        ref: str,
        value: Any,
        *,
        bold: bool = False,
        italic: bool = False,
        fill: str = "",
        font_color: str = "",
        size: float | None = None,
        align: str = "",
        number_format: str = "",
    ) -> str:
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

    def set_cells(
        self, sheet: str, rows: list[list], *, start_ref: str = "A1", has_header: bool = False
    ) -> str:
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
                ops.append(
                    {
                        "command": "set",
                        "path": f"/{sheet}/{ref}",
                        "props": props,
                    }
                )
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
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{sheet}/{ref}",
                "--prop",
                f"formula={formula}",
            ]
        )

    # ---- 行列格式 ----
    def set_column_width(self, sheet: str, col: str, width: float) -> str:
        """col 是列字母（如 'A'）或数字；width 是字符单位。"""
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{sheet}/col[{col}]",
                "--prop",
                f"width={width}",
            ]
        )

    def set_row_height(self, sheet: str, row: int, height: float) -> str:
        """row 是 1-based 行号；height 是磅。"""
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{sheet}/row[{row}]",
                "--prop",
                f"height={height}",
            ]
        )

    def autofit_column(self, sheet: str, col: str) -> str:
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{sheet}/col[{col}]",
                "--prop",
                "autofit=true",
            ]
        )

    def merge_cells(self, sheet: str, cell_range: str) -> str:
        """合并单元格。cell_range 如 'A1:C3'，锚点必须是左上格。"""
        anchor = cell_range.split(":")[0]
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{sheet}/{anchor}",
                "--prop",
                f"merge={cell_range}",
            ]
        )

    # ---- 图表 ----
    def add_chart(
        self,
        sheet: str,
        chart_type: str,
        data_range: str,
        *,
        categories: str = "",
        title: str = "",
        x: str = "2cm",
        y: str = "10cm",
        width: str = "15cm",
        height: str = "8cm",
        legend: str = "bottom",
    ) -> str:
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
            "add",
            self.doc_path,
            f"/{sheet}",
            "--type",
            "chart",
            "--prop",
            f"chartType={chart_type}",
            "--prop",
            f"dataRange={data_range}",
            "--prop",
            f"x={x}",
            "--prop",
            f"y={y}",
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
        return self.runner.run(
            [
                "set",
                self.doc_path,
                f"/{sheet}",
                "--prop",
                f"autoFilter={val}",
            ]
        )

    # ---- 条件格式 ----
    def add_conditional_format(self, sheet: str, cf_type: str, cell_range: str, **props) -> str:
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
            "add",
            self.doc_path,
            f"/{sheet}",
            "--type",
            cf_type,
            "--prop",
            f"sqref={cell_range}",
        ]
        for k, v in props.items():
            args += ["--prop", f"{k}={v}"]
        return self.runner.run(args)

    # ---- 透视表 ----
    def add_pivot_table(
        self,
        sheet: str,
        source: str,
        *,
        rows: str = "",
        cols: str = "",
        values: str = "",
        filters: str = "",
        position: str = "",
        name: str = "",
        layout: str = "",
    ) -> str:
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
            "add",
            self.doc_path,
            f"/{sheet}",
            "--type",
            "pivottable",
            "--prop",
            f"source={source}",
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
    def add_list_table(
        self,
        sheet: str,
        cell_range: str,
        *,
        name: str = "",
        style: str = "medium2",
        total_row: bool = False,
    ) -> str:
        """把单元格区域转成真正的 Excel 表格（带样式、筛选按钮、结构化引用）。

        参数:
            sheet: 工作表名。
            cell_range: 区域，如 'A1:C10'（首行需是表头）。
            name: 表格名。
            style: 表样式。常用 'medium1'~'medium4', 'light1'~'light3', 'dark1'~'dark2'。
            total_row: 是否显示汇总行。
        """
        args = [
            "add",
            self.doc_path,
            f"/{sheet}",
            "--type",
            "table",
            "--prop",
            f"ref={cell_range}",
            "--prop",
            f"style={style}",
        ]
        if name:
            args += ["--prop", f"name={name}"]
        if total_row:
            args += ["--prop", "totalRow=true"]
        return self.runner.run(args)

    # ---- 命名区域 ----
    def add_named_range(
        self, name: str, refers_to: str, *, scope: str = "workbook", comment: str = ""
    ) -> str:
        """定义命名区域（供公式用名称引用）。

        参数:
            name: 名称（如 'Revenue'）。
            refers_to: 引用，【不带前导 =】，如 'Sheet1!$A$1:$C$10'。
            scope: 作用域，'workbook' 或工作表名。
            comment: 备注（Name Manager 里显示）。
        """
        args = [
            "add",
            self.doc_path,
            "/",
            "--type",
            "namedrange",
            "--prop",
            f"name={name}",
            "--prop",
            f"ref={refers_to}",
            "--prop",
            f"scope={scope}",
        ]
        if comment:
            args += ["--prop", f"comment={comment}"]
        return self.runner.run(args)

    # ---- 数据验证 ----
    def add_validation(
        self,
        sheet: str,
        cell_range: str,
        val_type: str,
        *,
        formula1: str = "",
        formula2: str = "",
        operator: str = "",
        in_cell_dropdown: bool = True,
        prompt: str = "",
        error_msg: str = "",
    ) -> str:
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
            "add",
            self.doc_path,
            f"/{sheet}",
            "--type",
            "validation",
            "--prop",
            f"ref={cell_range}",
            "--prop",
            f"type={val_type}",
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
