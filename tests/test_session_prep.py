"""session.prep 纯函数测试。"""

from __future__ import annotations

import dataclasses

from office_agent.session.prep import (
    build_doc_path,
    official_header_fields,
    resolve_output_kind,
)


class TestResolveOutputKind:
    def test_official_forces_docx(self):
        kind, score = resolve_output_kind("写一份通知", doc_type="通知")
        assert kind == "docx"
        assert score > 0

    def test_excel_keywords(self):
        kind, score = resolve_output_kind("做一份销售数据 Excel 表格")
        assert kind == "xlsx"
        assert score > 0

    def test_ambiguous(self):
        kind, score = resolve_output_kind("随便弄点东西")
        assert kind is None
        assert score == 0


class TestBuildDocPath:
    def test_extension(self, tmp_path, monkeypatch):
        from office_agent import config
        from office_agent.session import prep

        new_settings = dataclasses.replace(config.settings, output_dir=tmp_path)
        monkeypatch.setattr(config, "settings", new_settings)
        monkeypatch.setattr(prep, "settings", new_settings)
        path = build_doc_path("项目周报", kind="docx")
        assert path.endswith(".docx")
        assert str(tmp_path) in path


class TestOfficialHeaderFields:
    def test_upward_has_signer(self):
        keys = [f["key"] for f in official_header_fields("报告")]
        assert "org" in keys
        assert "signer" in keys

    def test_downward_no_signer(self):
        keys = [f["key"] for f in official_header_fields("通知")]
        assert "org" in keys
        assert "signer" not in keys
