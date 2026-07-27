"""Excel（.xlsx）专属 LLM 工具。"""

from __future__ import annotations

import logging
from typing import Any

from langchain_core.tools import tool

from office_agent.officecli import (
    OfficeCLIError,
)
from office_agent.tools.session import (
    _tool,
    _wrong_kind_msg,
    session_doc_kind,
)

logger = logging.getLogger(__name__)


@tool
def add_sheet(name: str, tab_color: str = "") -> str:
    """【Excel 专属】添加一张工作表（tab）。

    新建的 xlsx 默认自带一张名为 'Sheet1' 的空工作表。本工具再加新的工作表。
    典型流程: create_doc → add_sheet('销售数据') → set_cells('销售数据', ...)。

    参数:
        name: 工作表名（如 '销售数据'）。后续 set_cell/set_cells/add_excel_chart
              都要传这个名字定位工作表。
        tab_color: 可选标签色，6位十六进制（如 '4472C4' 蓝色）。美化用。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_sheet", "xlsx", "docx 用 add_table；pptx 无工作表概念")
    try:
        return _tool().add_sheet(name, tab_color=tab_color)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"添加工作表失败: {e}"


@tool
def set_cell(
    sheet: str, ref: str, value: Any, bold: bool = False, fill: str = "", number_format: str = ""
) -> str:
    """【Excel 专属】写单个单元格。最基础、最灵活的写入工具。

    参数:
        sheet: 工作表名（如 'Sheet1' 或 '销售数据'）。
        ref: 单元格地址，列字母+行号，如 'A1' / 'B2' / 'AA10'。
        value: 单元格值。数字会存为数字（参与计算），字符串存为文本。
               注意：纯数字字符串如电话号 '01234' 想保留前导 0，
               要传 number_format='@' 强制文本。
        bold: 是否加粗（表头常用）。
        fill: 背景色，6位十六进制（如 'FFFF00' 黄色高亮）。
        number_format: 数字格式代码。常用:
            '@' 文本（保留前导0）
            '#,##0' 千分位整数
            '#,##0.00' 千分位两位小数
            '0%' 百分比
            'yyyy-mm-dd' 日期
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg(
            "set_cell", "xlsx", "docx 用 add_table/add_paragraph；pptx 用 add_textbox"
        )
    try:
        return _tool().set_cell(  # type: ignore[union-attr]
            sheet,
            ref,
            value,
            bold=bold,
            fill=fill,
            number_format=number_format,
        )
    except OfficeCLIError as e:
        return f"写入单元格失败: {e}"


@tool
def set_cells(sheet: str, data: list[list], start: str = "A1", has_header: bool = False) -> str:
    """【Excel 专属】批量写二维数据到工作表。写表格数据【首选】本工具，比逐个 set_cell 高效得多。

    参数:
        sheet: 目标工作表名。
        data: 二维数组，外层=行、内层=单元格。数字/字符串混合均可。
              例: [["月份","销售额"],["1月",12000],["2月",15000]]
        start: 起始单元格地址（默认 'A1'）。内部按行列递增自动算出每个 cell 的地址。
               如 start='B2' 则 data[0][0] 写到 B2、data[0][1] 写到 C2、data[1][0] 写到 B3...
        has_header: 第一行是否作为表头（加粗）。默认 false。

    一次调用写入一整片区域，适合写完整的表格/报表数据。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_cells", "xlsx", "docx 用 add_table；pptx 用 add_slide_table")
    try:
        return _tool().set_cells(
            sheet,
            data,
            start_ref=start,  # type: ignore[union-attr]
            has_header=has_header,
        )
    except OfficeCLIError as e:
        return f"批量写入失败: {e}"


@tool
def set_formula(sheet: str, ref: str, formula: str) -> str:
    """【Excel 专属】在单元格写公式（不带前导 =）。

    参数:
        sheet: 工作表名。
        ref: 单元格地址（如 'D2'）。
        formula: 公式表达式，【不要】带前导 '='。
              例: 'SUM(B2:B10)' / 'AVERAGE(C2:C5)' / 'B2-C2' / 'B2*C2*0.1'。

    公式里引用的单元格必须已先用 set_cell/set_cells 写入值，否则结果是 0 或错误。"""
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_formula", "xlsx", "公式只在 Excel 表格中可用")
    try:
        return _tool().set_formula(sheet, ref, formula)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"写公式失败: {e}"


@tool
def add_excel_chart(
    sheet: str, chart_type: str, data_range: str, title: str = "", categories: str = ""
) -> str:
    """【Excel 专属】在工作表上加图表（基于已写入的单元格数据）。

    前提: 数据必须已先用 set_cells 写入工作表，本工具只是引用这些单元格画图。

    参数:
        sheet: 图表要放在哪张工作表（图浮在该 sheet 上）。
        chart_type: 图表类型。常用:
            'column' 柱形图（默认推荐，对比类目数据）
            'bar'    条形图（横向柱状图）
            'line'   折线图（趋势）
            'pie'    饼图（占比）
            'doughnut' 环形图
            'area'   面积图
            'scatter' 散点图（相关性）
        data_range: 数据源区域，【必须带工作表名前缀】。
              格式 '工作表名!起始:结束'，如 '销售数据!B1:C4'。
              默认【首列当分类轴、其余列当数据系列】。
              若想让每列都是系列，请用 categories 参数单独指定分类轴。
        title: 图表标题（可空）。
        categories: 可选。分类轴数据，两种写法:
              - 区域引用: '销售数据!A2:A4'（推荐，引用已写入的类目标签）
              - 逗号分隔: 'Q1,Q2,Q3,Q4'

    示例: 数据写在 销售数据!A1:C4（A列月份/B列销售/C列成本），
          调 add_excel_chart('销售数据','column','销售数据!B1:C4',
                             title='季度销售vs成本', categories='销售数据!A2:A4')
          → 画出 B/C 两列为系列、A 列月份为分类的柱形图。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_excel_chart", "xlsx", "docx/pptx 的图表暂不支持数据区域引用")
    try:
        return _tool().add_chart(  # type: ignore[union-attr]
            sheet,
            chart_type,
            data_range,
            categories=categories,
            title=title,
        )
    except OfficeCLIError as e:
        return f"添加图表失败: {e}"


# ============================================================
# PowerPoint 专属工具（pptx）
# ============================================================


@tool
def sort_sheet(sheet: str, keys: str, has_header: bool = True) -> str:
    """【Excel 专属】对工作表数据排序（就地修改行顺序）。

    参数:
        sheet: 工作表名。
        keys: 排序键，格式 '列字母 [asc|desc]'，多键逗号分隔。
              例: 'B desc'（按 B 列降序）/ 'A asc, B desc'（先 A 升再 B 降）/ 'C'（C 列升序）。
        has_header: 首行是否为表头（不参与排序）。默认 true。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("sort_sheet", "xlsx", "排序只在 Excel 中可用")
    try:
        return _tool().sort(sheet, keys, has_header=has_header)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"排序失败: {e}"


@tool
def set_autofilter(sheet: str, cell_range: str = "") -> str:
    """【Excel 专属】开启自动筛选（表头出现下拉箭头，可筛选）。

    参数:
        sheet: 工作表名。
        cell_range: 筛选区域如 'A1:D10'。留空=自动识别已用区域。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_autofilter", "xlsx", "自动筛选只在 Excel 中可用")
    try:
        return _tool().set_autofilter(sheet, cell_range)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"开启筛选失败: {e}"


@tool
def highlight_cells(
    sheet: str, cell_range: str, operator: str, value: str, fill: str = "FFFF00"
) -> str:
    """【Excel 专属】条件格式：高亮满足条件的单元格（如大于某值）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'C2:C100'。
        operator: 比较运算符 'greaterThan'/'lessThan'/'equal'/'between' 等。
        value: 比较值。between 时用 '低值,高值'（此时会自动用 value2）。
        fill: 高亮背景色，6位十六进制（默认 'FFFF00' 黄色）。

    例: 高亮销售额>10000 的单元格
        highlight_cells('销售','C2:C100','greaterThan','10000','FF0000')
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("highlight_cells", "xlsx", "条件格式只在 Excel 中可用")
    try:
        props: dict = {"operator": operator, "fill": fill}
        if "," in value and operator in ("between", "notBetween"):
            v1, v2 = value.split(",", 1)
            props["value"] = v1.strip()
            props["value2"] = v2.strip()
        else:
            props["value"] = value
        return _tool().add_conditional_format(  # type: ignore[union-attr]
            sheet,
            "cellIs",
            cell_range,
            **props,
        )
    except OfficeCLIError as e:
        return f"添加条件格式失败: {e}"


@tool
def add_color_scale(
    sheet: str,
    cell_range: str,
    min_color: str = "F8696B",
    mid_color: str = "FFEB84",
    max_color: str = "63BE7B",
) -> str:
    """【Excel 专属】条件格式：3色渐变（红-黄-绿，数据热力图）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'C2:C100'。
        min_color: 最小值颜色（默认红 F8696B）。
        mid_color: 中间值颜色（默认黄 FFEB84）。
        max_color: 最大值颜色（默认绿 63BE7B）。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_color_scale", "xlsx", "条件格式只在 Excel 中可用")
    try:
        return _tool().add_conditional_format(  # type: ignore[union-attr]
            sheet,
            "colorScale",
            cell_range,
            minColor=min_color,
            midColor=mid_color,
            maxColor=max_color,
        )
    except OfficeCLIError as e:
        return f"添加色阶失败: {e}"


@tool
def add_data_bar(sheet: str, cell_range: str, color: str = "638EC6") -> str:
    """【Excel 专属】条件格式：数据条（单元格内显示横向条形）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'C2:C100'。
        color: 条形颜色（默认蓝 638EC6）。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_data_bar", "xlsx", "条件格式只在 Excel 中可用")
    try:
        return _tool().add_conditional_format(  # type: ignore[union-attr]
            sheet,
            "dataBar",
            cell_range,
            color=color,
        )
    except OfficeCLIError as e:
        return f"添加数据条失败: {e}"


@tool
def add_pivot_table(
    sheet: str,
    source: str,
    rows: str,
    values: str,
    cols: str = "",
    filters: str = "",
    position: str = "",
) -> str:
    """【Excel 专属】加数据透视表（汇总分析的利器）。

    参数:
        sheet: 透视表放在哪张工作表。
        source: 源数据区域，如 'Sheet1!A1:D100'（必须含表头行，区域要带工作表名前缀）。
        rows: 行字段，逗号分隔，如 '区域,产品'。
        values: 值字段及聚合方式，格式 '字段:agg'，多个逗号分隔。
                agg ∈ sum/avg/count/max/min。例: '销售额:sum,数量:count'。
        cols: 列字段，如 '季度'（可空）。
        filters: 筛选页字段，如 '年份'（可空）。
        position: 透视表左上角位置，如 'H1'。留空自动放在源数据旁。

    例: 按 region 汇总 sales 总和
        add_pivot_table('Sheet1','Sheet1!A1:D100','region','sales:sum',position='F1')
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_pivot_table", "xlsx", "透视表只在 Excel 中可用")
    try:
        return _tool().add_pivot_table(  # type: ignore[union-attr]
            sheet,
            source,
            rows=rows,
            cols=cols,
            values=values,
            filters=filters,
            position=position,
        )
    except OfficeCLIError as e:
        return f"添加透视表失败: {e}"


@tool
def add_list_table(
    sheet: str, cell_range: str, style: str = "medium2", total_row: bool = False
) -> str:
    """【Excel 专属】把单元格区域转成真正的 Excel 表格（带样式、筛选按钮、结构化引用）。

    与普通写数据的区别: 真 Excel 表格自动有蓝色条纹样式、表头筛选按钮、
    可用结构化引用（如 Table1[销售额]），还能一键汇总。

    参数:
        sheet: 工作表名。
        cell_range: 区域，如 'A1:C10'（首行需是表头）。
        style: 表样式，如 'medium2'(默认,蓝)/'medium4'(绿)/'light1'(浅)。
        total_row: 是否显示汇总行。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_list_table", "xlsx", "Excel 表格只在 Excel 中可用")
    try:
        return _tool().add_list_table(  # type: ignore[union-attr]
            sheet,
            cell_range,
            style=style,
            total_row=total_row,
        )
    except OfficeCLIError as e:
        return f"添加表格失败: {e}"


@tool
def add_dropdown(sheet: str, cell_range: str, options: str) -> str:
    """【Excel 专属】给单元格加下拉列表（数据验证）。

    参数:
        sheet: 工作表名。
        cell_range: 应用区域，如 'B2:B100'。
        options: 选项，逗号分隔，如 '是,否,待定'。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("add_dropdown", "xlsx", "下拉列表只在 Excel 中可用")
    try:
        return _tool().add_validation(  # type: ignore[union-attr]
            sheet,
            cell_range,
            "list",
            formula1=options,
            in_cell_dropdown=True,
        )
    except OfficeCLIError as e:
        return f"添加下拉列表失败: {e}"


@tool
def merge_cells(sheet: str, cell_range: str) -> str:
    """【Excel 专属】合并单元格。cell_range 如 'A1:D1'（左上格为锚点）。"""
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("merge_cells", "xlsx", "合并单元格只在 Excel 中可用")
    try:
        return _tool().merge_cells(sheet, cell_range)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"合并单元格失败: {e}"


@tool
def set_column_width(sheet: str, col: str, width: float) -> str:
    """【Excel 专属】设置列宽。

    参数:
        sheet: 工作表名。
        col: 列字母（如 'A'）。
        width: 宽度（字符单位，如 20）。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("set_column_width", "xlsx", "列宽只在 Excel 中可设")
    try:
        return _tool().set_column_width(sheet, col, width)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"设置列宽失败: {e}"


@tool
def autofit_column(sheet: str, col: str) -> str:
    """【Excel 专属】自动调整列宽以适应内容。

    参数:
        sheet: 工作表名。
        col: 列字母（如 'A'）。也可传 'A:C' 一次调整多列。
    """
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("autofit_column", "xlsx", "自动列宽只在 Excel 中可用")
    try:
        return _tool().autofit_column(sheet, col)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"自动列宽失败: {e}"


@tool
def rename_sheet(old_name: str, new_name: str) -> str:
    """【Excel 专属】重命名工作表。"""
    if session_doc_kind() != "xlsx":
        return _wrong_kind_msg("rename_sheet", "xlsx", "工作表只在 Excel 中有")
    try:
        return _tool().rename_sheet(old_name, new_name)  # type: ignore[union-attr]
    except OfficeCLIError as e:
        return f"重命名失败: {e}"


# ============================================================
# PowerPoint 进阶工具（pptx）—— 均与占位符模式兼容，不会重叠
# ============================================================
