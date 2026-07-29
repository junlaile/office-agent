"""目录（TOC）与 updateFields 集成测试（真调 officecli.exe）。

标记 @pytest.mark.integration，默认 skip。
显式运行：``uv run pytest -m integration -k toc``

验证：
  1. add_heading 写出的段落含 <w:outlineLvl>（让 Word 目录能收录标题）。
  2. add_toc 后 settings.xml 【不含】 updateFields=true（避免打开 Word
     弹「域可能引用了其它文件」提示）。
  3. 文档仍含 TOC 域，用户可手动更新。
"""

from __future__ import annotations

import re
import zipfile

import pytest

from office_agent.office.doc import DocTool

pytestmark = pytest.mark.integration


@pytest.fixture
def toc_doc(tmp_path):
    """真实创建一个 docx，加几级标题 + 目录，返回路径。"""
    path = str((tmp_path / "toc.docx").resolve())
    tool = DocTool(path)
    tool.create()
    tool.add_heading("一、概述", level=1)
    tool.add_paragraph("正文段落。")
    tool.add_heading("（一）背景", level=2)
    tool.add_paragraph("更多正文。")
    tool.add_heading("二、结论", level=1)
    tool.add_toc(title="目录")
    yield path
    try:
        tool.close()
    except Exception:  # noqa: BLE001
        pass


def _read_xml(path: str, part: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode("utf-8", "ignore")


def test_heading_paragraphs_have_outline_level(toc_doc):
    """add_heading 产出的段落含 <w:outlineLvl>，否则 TOC 收不到标题。"""
    xml = _read_xml(toc_doc, "word/document.xml")
    lvls = re.findall(r'<w:outlineLvl w:val="(\d+)"', xml)
    # 3 个标题：H1→0, H2→1, H1→0
    assert lvls == ["0", "1", "0"], f"outlineLvl 不符: {lvls}"


def test_toc_field_present(toc_doc):
    """文档含 TOC 域。"""
    xml = _read_xml(toc_doc, "word/document.xml")
    assert re.search(r'TOC\s*\\o', xml), "未找到 TOC 域指令"


def test_updatefields_cleared(toc_doc):
    """add_toc 后不应留下 updateFields=true，避免打开弹窗。"""
    xml = _read_xml(toc_doc, "word/settings.xml")
    assert not re.search(r"<w:updateFields\b", xml), (
        "settings.xml 仍含 updateFields，打开 Word 会弹域更新提示"
    )


def test_clear_updatefields_is_idempotent(toc_doc):
    """重复调用 _clear_updatefields 不破坏文档。"""
    tool = DocTool(toc_doc)
    tool._clear_updatefields()
    tool._clear_updatefields()
    xml = _read_xml(toc_doc, "word/settings.xml")
    assert not re.search(r"<w:updateFields\b", xml)
    doc = _read_xml(toc_doc, "word/document.xml")
    assert "outlineLvl" in doc
    assert re.search(r'TOC\s*\\o', doc)
