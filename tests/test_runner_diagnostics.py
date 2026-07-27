"""officecli runner 失败分类与诊断提示。"""

from __future__ import annotations

import subprocess
from unittest.mock import MagicMock

import pytest

from office_agent.office.runner import (
    OfficeCLIError,
    _Runner,
    classify_officecli_stderr,
)


class TestClassifyOfficecliStderr:
    def test_system_private_xml(self):
        stderr = (
            "Unhandled exception: System.IO.FileNotFoundException:\n"
            "File name: 'System.Private.Xml, Version=10.0.0.0'"
        )
        hint = classify_officecli_stderr(stderr)
        assert hint is not None
        assert ".NET" in hint
        assert "System.Private.Xml" in hint or "fetch_officecli" in hint

    def test_could_not_load_assembly(self):
        hint = classify_officecli_stderr(
            "Could not load file or assembly 'System.Private.CoreLib'"
        )
        assert hint is not None
        assert "fetch_officecli" in hint

    def test_unrelated_stderr_returns_none(self):
        assert classify_officecli_stderr("path not found: /body/p[9]") is None

    def test_plain_filenotfound_without_assembly_meta_returns_none(self):
        assert classify_officecli_stderr(
            "Unhandled exception: System.IO.FileNotFoundException: doc missing"
        ) is None

    def test_empty_returns_none(self):
        assert classify_officecli_stderr("") is None
        assert classify_officecli_stderr(None) is None


class TestRunnerDotnetDiagnostics:
    def test_nonzero_includes_dotnet_hint(self, monkeypatch, tmp_path):
        fake_bin = tmp_path / "officecli"
        fake_bin.write_text("#!/bin/sh\n")
        fake_bin.chmod(0o755)

        proc = MagicMock()
        proc.returncode = 1
        proc.stdout = ""
        proc.stderr = (
            "Unhandled exception: System.IO.FileNotFoundException: "
            "File name: 'System.Private.Xml, Version=10.0.0.0'"
        )
        monkeypatch.setattr(subprocess, "run", MagicMock(return_value=proc))

        runner = _Runner(bin_path=str(fake_bin))
        with pytest.raises(OfficeCLIError) as ei:
            runner.run(["add", "a.docx", "/body", "--type", "paragraph"])
        msg = str(ei.value)
        assert "诊断:" in msg
        assert ".NET" in msg
        assert ei.value.returncode == 1
        assert "System.Private.Xml" in (ei.value.stderr or "")

    def test_file_not_found_message(self, monkeypatch, tmp_path):
        fake_bin = tmp_path / "missing-officecli"
        monkeypatch.setattr(
            subprocess,
            "run",
            MagicMock(side_effect=FileNotFoundError("No such file")),
        )
        runner = _Runner(bin_path=str(fake_bin))
        with pytest.raises(OfficeCLIError, match="找不到 officecli") as ei:
            runner.run(["create", "a.docx"])
        assert "fetch_officecli" in str(ei.value)
