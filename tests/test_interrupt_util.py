"""pending_interrupt 纯函数测试。"""

from __future__ import annotations

from types import SimpleNamespace

from office_agent.session.interrupt_util import pending_interrupt


class TestPendingInterrupt:
    def test_none_snapshot(self):
        graph = SimpleNamespace(get_state=lambda config: None)
        assert pending_interrupt(graph, {}) is None

    def test_no_tasks(self):
        graph = SimpleNamespace(
            get_state=lambda config: SimpleNamespace(tasks=[])
        )
        assert pending_interrupt(graph, {}) is None

    def test_dict_payload(self):
        intr = SimpleNamespace(value={"title": "q", "fields": []})
        task = SimpleNamespace(interrupts=[intr])
        graph = SimpleNamespace(
            get_state=lambda config: SimpleNamespace(tasks=[task])
        )
        assert pending_interrupt(graph, {}) == {"title": "q", "fields": []}

    def test_str_payload(self):
        intr = SimpleNamespace(value="hello")
        task = SimpleNamespace(interrupts=[intr])
        graph = SimpleNamespace(
            get_state=lambda config: SimpleNamespace(tasks=[task])
        )
        assert pending_interrupt(graph, {}) == {"question": "hello"}

    def test_none_value(self):
        intr = SimpleNamespace(value=None)
        task = SimpleNamespace(interrupts=[intr])
        graph = SimpleNamespace(
            get_state=lambda config: SimpleNamespace(tasks=[task])
        )
        assert pending_interrupt(graph, {}) == {}
