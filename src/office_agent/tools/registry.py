"""工具元数据与注册表。

工具函数只负责定义参数和执行逻辑；可用文档类型、交互方式和副作用等策略
集中保存在本模块，供 LLM 绑定和 LangGraph 执行节点统一决策。
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from types import MappingProxyType
from typing import TYPE_CHECKING, Mapping, Sequence

if TYPE_CHECKING:
    from langchain_core.tools import BaseTool


class ExecutionMode(StrEnum):
    """工具进入哪种执行流程。"""

    DIRECT = "direct"
    INTERACTION = "interaction"
    CONFIRMATION = "confirmation"


class SideEffect(StrEnum):
    """工具副作用分类，供确认、审计和重放策略使用。"""

    NONE = "none"
    READ = "read"
    WRITE = "write"
    INIT = "init"
    HUMAN = "human"
    TERMINAL = "terminal"


DOCX_KINDS = frozenset({"docx"})
XLSX_KINDS = frozenset({"xlsx"})
PPTX_KINDS = frozenset({"pptx"})
ALL_DOCUMENT_KINDS = frozenset({"docx", "xlsx", "pptx"})


def document_kind_from_path(doc_path: str) -> str:
    """从输出路径推断文档类型；未知扩展名与会话路由一致，默认 Word。"""
    path = doc_path.lower()
    if path.endswith(".xlsx"):
        return "xlsx"
    if path.endswith(".pptx"):
        return "pptx"
    return "docx"


@dataclass(frozen=True, slots=True)
class ToolSpec:
    """一个工具及其执行策略。"""

    tool: BaseTool
    document_kinds: frozenset[str]
    execution_mode: ExecutionMode = ExecutionMode.DIRECT
    side_effect: SideEffect = SideEffect.WRITE
    can_batch: bool = True
    confirmation_title: str = ""

    @property
    def name(self) -> str:
        return self.tool.name


class ToolRegistry:
    """不可变工具注册表。"""

    def __init__(self, specs: Sequence[ToolSpec]) -> None:
        ordered = tuple(specs)
        names = [spec.name for spec in ordered]
        if len(names) != len(set(names)):
            raise ValueError("工具注册表中存在重复名称")

        for spec in ordered:
            unknown = spec.document_kinds - ALL_DOCUMENT_KINDS
            if not spec.document_kinds or unknown:
                raise ValueError(
                    f"工具 {spec.name} 的 document_kinds 无效: {sorted(spec.document_kinds)}"
                )
            if spec.execution_mode is not ExecutionMode.DIRECT and spec.can_batch:
                raise ValueError(f"交互/确认工具 {spec.name} 必须设置 can_batch=False")
            if (
                spec.execution_mode is ExecutionMode.CONFIRMATION
                and spec.side_effect in {SideEffect.NONE, SideEffect.READ, SideEffect.HUMAN}
            ):
                raise ValueError(f"确认工具 {spec.name} 必须声明实际副作用")

        self._specs = ordered
        self._spec_by_name: Mapping[str, ToolSpec] = MappingProxyType(
            {spec.name: spec for spec in ordered}
        )

    @property
    def specs(self) -> tuple[ToolSpec, ...]:
        return self._specs

    @property
    def all_tools(self) -> list[BaseTool]:
        return [spec.tool for spec in self._specs]

    @property
    def tool_by_name(self) -> dict[str, BaseTool]:
        return {name: spec.tool for name, spec in self._spec_by_name.items()}

    @property
    def spec_by_name(self) -> Mapping[str, ToolSpec]:
        return self._spec_by_name

    def get(self, name: str | None) -> ToolSpec | None:
        return self._spec_by_name.get(name or "")

    def bindable_tools(self, doc_path: str) -> list[BaseTool]:
        kind = document_kind_from_path(doc_path)
        return [spec.tool for spec in self._specs if kind in spec.document_kinds]
