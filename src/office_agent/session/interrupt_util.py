"""从 LangGraph checkpoint 读取 pending interrupt（无 UI）。"""

from __future__ import annotations

from typing import Any


def pending_interrupt(graph: Any, config: dict) -> dict[str, Any] | None:
    """若当前挂在 interrupt，返回 payload dict；否则 None。"""
    snapshot = graph.get_state(config)
    if snapshot is None:
        return None
    # 遍历所有待执行任务，收集其挂起的 interrupt
    interrupts: list = []
    for t in snapshot.tasks or []:
        if hasattr(t, "interrupts") and t.interrupts:
            interrupts.extend(t.interrupts)
    if not interrupts:
        return None
    value = interrupts[0].value
    if value is None:
        return {}
    if isinstance(value, dict):
        return value
    # 非 dict 的 payload（如纯字符串问题）统一包装成 question 字段
    return {"question": str(value)}
