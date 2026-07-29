"""DocTool._clear_updatefields 单测：用最小 docx zip 覆盖清除逻辑。"""

from __future__ import annotations

import zipfile

from office_agent.office.doc import DocTool


def _make_docx(path: str, *, settings_xml: str, extra: dict[str, str] | None = None) -> None:
    """写一个最小可被 ZipFile 打开的伪 docx。"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        z.writestr("word/settings.xml", settings_xml)
        z.writestr("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>")
        if extra:
            for name, content in extra.items():
                z.writestr(name, content)


def _read_settings(path: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/settings.xml").decode("utf-8")


class TestClearUpdatefields:
    def test_removes_val_true(self, tmp_path):
        path = str(tmp_path / "a.docx")
        _make_docx(
            path,
            settings_xml=(
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:updateFields w:val="true"/>'
                "<w:zoom w:percent=\"100\"/>"
                "</w:settings>"
            ),
        )
        DocTool(path)._clear_updatefields()
        xml = _read_settings(path)
        assert "<w:updateFields" not in xml
        assert "<w:zoom" in xml  # 其它设置保留

    def test_removes_empty_tag(self, tmp_path):
        path = str(tmp_path / "b.docx")
        _make_docx(
            path,
            settings_xml=(
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:updateFields />"
                "</w:settings>"
            ),
        )
        DocTool(path)._clear_updatefields()
        assert "<w:updateFields" not in _read_settings(path)

    def test_noop_when_absent(self, tmp_path):
        path = str(tmp_path / "c.docx")
        original = (
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:zoom w:percent=\"100\"/>"
            "</w:settings>"
        )
        _make_docx(path, settings_xml=original)
        DocTool(path)._clear_updatefields()
        assert _read_settings(path) == original

    def test_idempotent(self, tmp_path):
        path = str(tmp_path / "d.docx")
        _make_docx(
            path,
            settings_xml=(
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:updateFields w:val="true"/>'
                "</w:settings>"
            ),
        )
        tool = DocTool(path)
        tool._clear_updatefields()
        tool._clear_updatefields()
        assert "<w:updateFields" not in _read_settings(path)
        # document 仍在
        with zipfile.ZipFile(path) as z:
            assert "word/document.xml" in z.namelist()
