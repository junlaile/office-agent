"""officecli 封装的纯函数测试：staticmethod + 列字母/ref 工具。

这些是 DocTool/ExcelTool 的辅助方法，纯函数无外部依赖。
真实 subprocess 调用的测试见 test_officecli_argv.py（用 FakeRunner）。
"""

from __future__ import annotations

import pytest

from office_agent.doc_tool import DocTool
from office_agent.excel_tool import _col_to_letter, _ref_at


# ============================================================
# DocTool._parse_tbl_index: 从 add table 输出解析表索引
# ============================================================
class TestParseTblIndex:
    def test_normal(self):
        assert DocTool._parse_tbl_index("Added table at /body/tbl[3]") == 3

    def test_first_table(self):
        assert DocTool._parse_tbl_index("Added table at /body/tbl[1]") == 1

    def test_large_index(self):
        assert DocTool._parse_tbl_index("/body/tbl[99]") == 99

    def test_no_match_returns_zero(self):
        assert DocTool._parse_tbl_index("Added paragraph at /body/p[1]") == 0

    def test_empty_returns_zero(self):
        assert DocTool._parse_tbl_index("") == 0

    def test_none_returns_zero(self):
        assert DocTool._parse_tbl_index(None) == 0  # type: ignore[arg-type]


# ============================================================
# DocTool._build_table_ops: 构造写单元格的 batch ops
# ============================================================
class TestBuildTableOps:
    def test_basic_2x2(self):
        """2行2列 → 4 个 set op。"""
        data = [["a", "b"], ["c", "d"]]
        ops = DocTool._build_table_ops(data, tbl_index=1, has_header=True)
        assert len(ops) == 4
        assert all(op["command"] == "set" for op in ops)

    def test_paths_correct(self):
        """路径格式 /body/tbl[N]/tr[R]/tc[C]。"""
        data = [["x"]]
        ops = DocTool._build_table_ops(data, tbl_index=2, has_header=False)
        assert ops[0]["path"] == "/body/tbl[2]/tr[1]/tc[1]"

    def test_header_bold(self):
        """has_header=True 时表头行加 bold。"""
        data = [["h1", "h2"], ["v1", "v2"]]
        ops = DocTool._build_table_ops(data, tbl_index=1, has_header=True)
        # 前 2 个 op（表头）应有 bold=true
        assert ops[0]["props"].get("bold") == "true"
        assert ops[1]["props"].get("bold") == "true"
        # 数据行无 bold
        assert "bold" not in ops[2]["props"]

    def test_no_header_no_bold(self):
        """has_header=False 无 bold。"""
        data = [["a"]]
        ops = DocTool._build_table_ops(data, tbl_index=1, has_header=False)
        assert "bold" not in ops[0]["props"]

    def test_text_values(self):
        """单元格值作为 text 写入 props。"""
        data = [["hello", 123]]
        ops = DocTool._build_table_ops(data, tbl_index=1, has_header=False)
        assert ops[0]["props"]["text"] == "hello"
        assert ops[1]["props"]["text"] == "123"

    def test_zero_index_defaults_to_1(self):
        """tbl_index<=0 时回退到 1。"""
        data = [["a"]]
        ops = DocTool._build_table_ops(data, tbl_index=0, has_header=False)
        assert "tbl[1]" in ops[0]["path"]


# ============================================================
# _col_to_letter / _ref_at: Excel 列字母工具
# ============================================================
class TestColToLetter:
    @pytest.mark.parametrize(
        "n,expected",
        [
            (1, "A"),
            (2, "B"),
            (26, "Z"),
            (27, "AA"),
            (28, "AB"),
            (52, "AZ"),
            (53, "BA"),
            (702, "ZZ"),
            (703, "AAA"),
        ],
    )
    def test_conversion(self, n, expected):
        assert _col_to_letter(n) == expected

    def test_zero_or_negative_clamps_to_1(self):
        """<=0 钳制到 1。"""
        assert _col_to_letter(0) == "A"
        assert _col_to_letter(-5) == "A"


class TestRefAt:
    @pytest.mark.parametrize(
        "row,col,expected",
        [
            (1, 1, "A1"),
            (2, 3, "C2"),
            (10, 27, "AA10"),
            (100, 26, "Z100"),
        ],
    )
    def test_conversion(self, row, col, expected):
        assert _ref_at(row, col) == expected
