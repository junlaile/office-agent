"""交互适配器测试。"""

from office_agent.cli.interactions import CLIInteractionAdapter


def test_form_response_is_wrapped_with_request_id():
    adapter = CLIInteractionAdapter(
        collect_form=lambda _request: {"name": "张三"},
        collect_question=lambda _request: "unused",
    )
    result = adapter.collect(
        {
            "request_id": "r1",
            "kind": "form",
            "fields": [{"key": "name", "label": "姓名"}],
        }
    )
    assert result == {"request_id": "r1", "answers": {"name": "张三"}}


def test_confirmation_maps_selected_value_to_accepted():
    adapter = CLIInteractionAdapter(
        collect_form=lambda _request: {},
        collect_question=lambda _request: "确认",
    )
    result = adapter.collect(
        {
            "request_id": "r2",
            "kind": "confirmation",
            "question": "继续？",
            "options": ["确认", "取消"],
        }
    )
    assert result["request_id"] == "r2"
    assert result["accepted"] is True


def test_legacy_question_remains_supported():
    adapter = CLIInteractionAdapter(
        collect_form=lambda _request: {},
        collect_question=lambda _request: "自定义答案",
    )
    result = adapter.collect({"question": "请回答"})
    assert result == {"request_id": "", "value": "自定义答案"}
