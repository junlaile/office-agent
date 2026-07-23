"""公文模板 merge 集成测试（真调 officecli + 真实模板）。

标记 @pytest.mark.integration，默认 skip。
显式运行：``uv run pytest -m integration -k official_merge``
"""

from __future__ import annotations

from pathlib import Path

import pytest

from office_agent.officecli import merge_template
from office_agent.domain.templates import OFFICIAL_DOCS, default_merge_data, template_path

pytestmark = pytest.mark.integration


@pytest.fixture
def merged_output(tmp_path):
    """对通知模板做真实 merge，返回输出路径。"""
    out = str((tmp_path / "merged.docx").resolve())
    tmpl = str(template_path("通知"))
    data = default_merge_data("通知", org="市公安局", doc_no="X公发〔2026〕1号")
    merge_template(tmpl, out, data)
    return out


def test_merge_creates_file(merged_output):
    """merge 后文件存在且非空。"""
    p = Path(merged_output)
    assert p.exists()
    assert p.stat().st_size > 0


def test_merge_no_residual_placeholders(merged_output):
    """merge 后无 {{key}} 残留。"""
    # 用 view_text 读（需 DocTool）
    from office_agent.office.doc import DocTool

    text = DocTool(merged_output).view_text()
    assert "{{" not in text
    assert "}}" not in text


def test_merge_filled_org(merged_output):
    """org 槽位被真实值替换。"""
    from office_agent.office.doc import DocTool

    text = DocTool(merged_output).view_text()
    assert "市公安局" in text


def test_all_15_templates_merge(tmp_path):
    """所有 15 个模板都能被 merge（批量验证）。"""
    for d in OFFICIAL_DOCS:
        out = str((tmp_path / f"test_{d.name}.docx").resolve())
        data = default_merge_data(d.name, org="测试机关")
        tmpl = str(template_path(d.name))
        # 不应抛错
        merge_template(tmpl, out, data)
        assert Path(out).exists()


def test_merged_passes_validation(merged_output):
    """merge 产物通过 OpenXML 校验。"""
    from office_agent.office.doc import DocTool

    result = DocTool(merged_output).validate()
    assert "passed" in result.lower() or "no errors" in result.lower()
