"""DocTool._ensure_updatefields 单测：用最小 docx zip 覆盖补丁逻辑。"""

from __future__ import annotations

import re
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
        z.writestr(
            "word/document.xml",
            "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>",
        )
        if extra:
            for name, content in extra.items():
                z.writestr(name, content)


def _read_settings(path: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/settings.xml").decode("utf-8")


class TestEnsureUpdatefields:
    def test_patches_empty_tag(self, tmp_path):
        path = str(tmp_path / "a.docx")
        _make_docx(
            path,
            settings_xml=(
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:updateFields />"
                '<w:zoom w:percent="100"/>'
                "</w:settings>"
            ),
        )
        DocTool(path)._ensure_updatefields()
        xml = _read_settings(path)
        assert re.search(r'<w:updateFields\s+w:val="true"\s*/?>', xml)
        assert "<w:zoom" in xml

    def test_inserts_when_absent(self, tmp_path):
        path = str(tmp_path / "b.docx")
        _make_docx(
            path,
            settings_xml=(
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:zoom w:percent="100"/>'
                "</w:settings>"
            ),
        )
        DocTool(path)._ensure_updatefields()
        xml = _read_settings(path)
        assert re.search(r'<w:updateFields\s+w:val="true"\s*/?>', xml)
        assert "<w:zoom" in xml

    def test_noop_when_already_true(self, tmp_path):
        path = str(tmp_path / "c.docx")
        original = (
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:updateFields w:val="true"/>'
            "</w:settings>"
        )
        _make_docx(path, settings_xml=original)
        DocTool(path)._ensure_updatefields()
        assert _read_settings(path) == original

    def test_idempotent(self, tmp_path):
        path = str(tmp_path / "d.docx")
        _make_docx(
            path,
            settings_xml=(
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:updateFields />"
                "</w:settings>"
            ),
        )
        tool = DocTool(path)
        tool._ensure_updatefields()
        tool._ensure_updatefields()
        xml = _read_settings(path)
        assert len(re.findall(r"<w:updateFields", xml)) == 1
        assert re.search(r'<w:updateFields\s+w:val="true"\s*/?>', xml)
        with zipfile.ZipFile(path) as z:
            assert "word/document.xml" in z.namelist()
