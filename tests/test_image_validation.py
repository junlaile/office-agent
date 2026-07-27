"""图片来源预校验单测：_validate_image_source 各类源的处理。

不碰 officecli、不发真实网络请求（HTTP 用 monkeypatch mock urlopen）。
验证"图片不可用时不嵌入"的核心拦截逻辑。
"""

from __future__ import annotations

import io
from urllib.error import HTTPError, URLError

from office_agent.tools.common import _validate_image_source


class TestValidateLocalAndData:
    """本地路径与 data URI：纯文件系统/字符串判断，无网络。"""

    def test_empty_string_blocked(self):
        assert _validate_image_source("") == "图片来源为空"

    def test_whitespace_only_blocked(self):
        assert _validate_image_source("   ") == "图片来源为空"

    def test_data_uri_always_valid(self):
        # data URI 内联数据，总是可用
        assert _validate_image_source("data:image/png;base64,iVBORw0KGgo=") is None

    def test_data_uri_text_valid(self):
        assert _validate_image_source("data:text/plain,hello") is None

    def test_nonexistent_local_path_blocked(self, tmp_path):
        missing = str(tmp_path / "no_such_file.png")
        reason = _validate_image_source(missing)
        assert reason is not None
        assert "不存在" in reason
        assert missing in reason

    def test_existing_local_path_valid(self, tmp_path):
        img = tmp_path / "photo.png"
        img.write_bytes(b"\x89PNG fake")
        assert _validate_image_source(str(img)) is None


# ============================================================
# HTTP/HTTPS：mock urlopen 模拟各类响应
# ============================================================
class _FakeResponse:
    """假 HTTP 响应（context manager）。"""

    def __init__(self, status: int):
        self.status = status

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def _patch_urlopen(monkeypatch, side_effect):
    """把 common.urlopen 替换为返回 side_effect 的可调用对象。

    side_effect 可以是：状态码 int（包成 _FakeResponse）、或异常实例、或异常类。
    """
    from office_agent.tools import common

    def fake_urlopen(req, timeout=None):  # noqa: ARG001
        if isinstance(side_effect, BaseException):
            raise side_effect
        if isinstance(side_effect, type) and issubclass(side_effect, BaseException):
            raise side_effect("boom")
        return _FakeResponse(side_effect)

    monkeypatch.setattr(common, "urlopen", fake_urlopen)


class TestValidateHttp:
    """HTTP/HTTPS HEAD 探测的各种情形。"""

    def test_http_200_valid(self, monkeypatch):
        _patch_urlopen(monkeypatch, 200)
        assert _validate_image_source("https://example.com/a.png") is None

    def test_http_301_redirect_valid(self, monkeypatch):
        # 3xx 重定向 → 放行
        _patch_urlopen(monkeypatch, 301)
        assert _validate_image_source("https://example.com/a.png") is None

    def test_http_404_blocked(self, monkeypatch):
        _patch_urlopen(monkeypatch, 404)
        reason = _validate_image_source("https://example.com/missing.png")
        assert reason is not None
        assert "404" in reason
        assert "不存在" in reason

    def test_http_410_gone_blocked(self, monkeypatch):
        _patch_urlopen(monkeypatch, 410)
        reason = _validate_image_source("https://example.com/gone.png")
        assert reason is not None
        assert "410" in reason

    def test_http_405_method_not_allowed_passes(self, monkeypatch):
        # 405 = HEAD 不被支持 → 放行交给 officecli（不误拦）
        _patch_urlopen(monkeypatch, 405)
        assert _validate_image_source("https://example.com/a.png") is None

    def test_http_404_via_http_error_blocked(self, monkeypatch):
        # urlopen 抛 HTTPError(404) 也是 URLError 子类
        err = HTTPError("url", 404, "Not Found", {}, io.BytesIO(b""))
        _patch_urlopen(monkeypatch, err)
        reason = _validate_image_source("https://example.com/x.png")
        assert reason is not None
        assert "404" in reason

    def test_http_405_via_http_error_passes(self, monkeypatch):
        err = HTTPError("url", 405, "Method Not Allowed", {}, io.BytesIO(b""))
        _patch_urlopen(monkeypatch, err)
        assert _validate_image_source("https://example.com/x.png") is None

    def test_http_403_via_http_error_passes(self, monkeypatch):
        # 403 常见于图片 CDN 拒绝 HEAD/无 Referer 探测，但实际 GET 能取
        # → 放行交给 officecli（宁可放行，不可误拦）
        err = HTTPError("url", 403, "Forbidden", {}, io.BytesIO(b""))
        _patch_urlopen(monkeypatch, err)
        assert _validate_image_source("https://example.com/x.png") is None

    def test_head_probe_sends_user_agent(self, monkeypatch):
        # 无 UA 的探测请求常被 CDN 直接 403 → 必须带常规 UA
        captured = {}

        from office_agent.tools import common

        def fake_urlopen(req, timeout=None):  # noqa: ARG001
            captured["ua"] = req.get_header("User-agent")
            return _FakeResponse(200)

        monkeypatch.setattr(common, "urlopen", fake_urlopen)
        assert _validate_image_source("https://example.com/a.png") is None
        assert captured["ua"] and "Mozilla" in captured["ua"]

    def test_connection_error_blocked(self, monkeypatch):
        # DNS 失败/拒绝连接/超时 → 拦
        err = URLError("connection refused")
        _patch_urlopen(monkeypatch, err)
        reason = _validate_image_source("https://no-such-host-xxx.invalid/a.png")
        assert reason is not None
        assert "不可访问" in reason

    def test_timeout_blocked(self, monkeypatch):

        err = URLError(TimeoutError("timed out"))
        _patch_urlopen(monkeypatch, err)
        reason = _validate_image_source("https://slow-host.invalid/a.png")
        assert reason is not None
        assert "不可访问" in reason

    def test_unknown_exception_passes(self, monkeypatch):
        # 未知异常 → 放行（宁可放行不误拦，交给 officecli 兜底）
        _patch_urlopen(monkeypatch, ValueError("weird"))
        assert _validate_image_source("https://example.com/a.png") is None
