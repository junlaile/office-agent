"""公文模式编排测试：_prepare_official_doc 编排逻辑（mock merge_template）。

验证 main 流程里的公文预创建：识别 → merge → 预读模板正文 → 打印 → 返回
(文种名, 模板正文)。不真调 officecli.exe。
"""

from __future__ import annotations

import pytest

from office_agent import cli_ui
from office_agent.cli_ui import _prepare_official_doc


@pytest.fixture
def stub_doctool_viewtext(monkeypatch):
    """把 cli_ui.DocTool 换成假实现：view_text 返回固定正文，不碰 officecli。

    _prepare_official_doc 在 merge 后会调 DocTool(doc_path).view_text() 预读
    模板正文；测试里必须 stub 掉，否则会真调 officecli.exe。
    """

    class _FakeDocTool:
        def __init__(self, doc_path):
            self.doc_path = doc_path

        def view_text(self):
            return "[/body/p[4]] 范例标题\n[/body/p[5]] 范例主送\n[/body/p[6]] 正文范例"

    monkeypatch.setattr(cli_ui, "DocTool", _FakeDocTool)
    return _FakeDocTool


class TestPrepareOfficialDoc:
    """_prepare_official_doc 编排（mock merge + view_text）。"""

    def test_calls_merge_and_returns_type(self, monkeypatch, tmp_path, capsys, stub_doctool_viewtext):
        """正常情况：调 merge + 预读正文 + 返回 (文种名, 正文)。"""
        merge_calls = []

        def fake_merge(tmpl, out, data):
            merge_calls.append((tmpl, out, data))
            return "Merged OK"

        monkeypatch.setattr(cli_ui, "merge_template", fake_merge)
        out_path = str(tmp_path / "通知.docx")

        result = _prepare_official_doc("通知", out_path)
        assert result == ("通知", "[/body/p[4]] 范例标题\n[/body/p[5]] 范例主送\n[/body/p[6]] 正文范例")
        assert len(merge_calls) == 1
        tmpl, out, data = merge_calls[0]
        assert out == out_path
        assert "08-通知" in tmpl  # 模板路径含文件名
        # merge_data 无 {{ 残留
        assert all("{{" not in v for v in data.values())
        # 打印了提示
        captured = capsys.readouterr()
        assert "GB/T 9704" in captured.out or "模板创建" in captured.out

    def test_missing_template_returns_none(self, monkeypatch, tmp_path, stub_doctool_viewtext):
        """模板缺失时返回 (None, '')。"""
        nonexistent = tmp_path / "不存在.docx"
        # _prepare_official_doc 用的是 cli_ui 模块绑定的 template_path
        monkeypatch.setattr(cli_ui, "template_path", lambda dt: nonexistent)
        result = _prepare_official_doc("通知", str(tmp_path / "out.docx"))
        assert result == (None, "")

    def test_merge_failure_returns_none(self, monkeypatch, tmp_path, stub_doctool_viewtext):
        """merge 失败时返回 (None, '')（不抛）。"""
        from office_agent.cli_runner import OfficeCLIError

        def failing_merge(*a, **k):
            raise OfficeCLIError("merge 失败")

        monkeypatch.setattr(cli_ui, "merge_template", failing_merge)
        result = _prepare_official_doc("通知", str(tmp_path / "out.docx"))
        assert result == (None, "")

    def test_viewtext_failure_returns_empty_text(self, monkeypatch, tmp_path, stub_doctool_viewtext):
        """view_text 预读失败时不阻断：文种名照常返回，正文退化为空串。"""
        from office_agent.cli_runner import OfficeCLIError

        class _BoomDocTool:
            def __init__(self, doc_path):
                pass

            def view_text(self):
                raise OfficeCLIError("view 失败")

        monkeypatch.setattr(cli_ui, "merge_template", lambda *a, **k: "OK")
        monkeypatch.setattr(cli_ui, "DocTool", _BoomDocTool)
        result = _prepare_official_doc("通知", str(tmp_path / "out.docx"))
        assert result == ("通知", "")  # 预读失败优雅降级

    def test_merge_data_has_required_keys(self, monkeypatch, tmp_path, stub_doctool_viewtext):
        """传给 merge 的 data 含必要槽位。"""
        captured_data = {}

        def spy_merge(tmpl, out, data):
            captured_data.update(data)
            return "OK"

        monkeypatch.setattr(cli_ui, "merge_template", spy_merge)
        _prepare_official_doc("通知", str(tmp_path / "out.docx"))
        required = {"org", "doc_no", "date_cn", "issuer"}
        assert required <= set(captured_data.keys())

    def test_upward_doc_has_signer_in_data(self, monkeypatch, tmp_path, stub_doctool_viewtext):
        """上行文（请示）的 merge_data 含非空 signer。"""
        captured_data = {}

        def spy_merge(tmpl, out, data):
            captured_data.update(data)
            return "OK"

        monkeypatch.setattr(cli_ui, "merge_template", spy_merge)
        _prepare_official_doc("请示", str(tmp_path / "out.docx"))
        assert captured_data["signer"]  # 非空
