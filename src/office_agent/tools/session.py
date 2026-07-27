"""会话级文档路径基础设施（tools 包内共享）。

main.py 在启动 agent 前调用 ``set_session_doc(path)``；所有 @tool 工具内部
读取该路径并按扩展名路由到 DocTool / ExcelTool / PptxTool，LLM 无需传路径。

存储用 ContextVar + 模块级默认值双轨（与 ``cli.user_input`` 的 bridge 同一模式）:
    - ContextVar 支持同进程多上下文并行会话（如未来的服务化场景）；
    - LangGraph 可能在工作线程里执行节点（新线程读不到调用方的 ContextVar），
      此时回退到模块级默认值，行为与旧的模块全局一致。

本模块独立成文件（而非放在 ``tools/__init__.py``）是为了消除子模块与包
``__init__`` 之间的循环 import：doc/excel/pptx/common 都从这里 import 会话
基础设施，``__init__.py`` 再聚合它们，依赖方向单一。
"""

from __future__ import annotations

import logging
from contextvars import ContextVar

from office_agent.office.doc import DocTool
from office_agent.office.excel import ExcelTool
from office_agent.office.pptx import PptxTool
from office_agent.office.runner import OfficeCLIError

logger = logging.getLogger(__name__)

_doc_path_var: ContextVar[str | None] = ContextVar("session_doc_path", default=None)
_default_doc_path: str | None = None


def set_session_doc(path: str | None) -> None:
    """main.py 在启动 agent 前调用，设定本会话的文档路径。

    传 None 清空（供测试 teardown）。
    """
    global _default_doc_path
    _default_doc_path = path
    _doc_path_var.set(path)


def session_doc_path() -> str | None:
    return _doc_path_var.get() or _default_doc_path


def session_doc_kind() -> str:
    """返回当前会话的文档类型: 'docx' | 'xlsx' | 'pptx'。

    路径未初始化时抛错；扩展名无法识别时默认 'docx'。
    """
    path = session_doc_path()
    if not path:
        raise OfficeCLIError("会话文档路径未初始化（请先调用 set_session_doc）")
    p = path.lower()
    if p.endswith(".xlsx"):
        return "xlsx"
    if p.endswith(".pptx"):
        return "pptx"
    return "docx"


def _tool() -> DocTool | ExcelTool | PptxTool:
    """工厂：按当前会话扩展名返回对应的 Tool 实例。"""
    path = session_doc_path()
    if not path:
        raise OfficeCLIError("会话文档路径未初始化（请先调用 set_session_doc）")
    kind = session_doc_kind()
    if kind == "xlsx":
        return ExcelTool(path)
    if kind == "pptx":
        return PptxTool(path)
    return DocTool(path)


def _doc() -> DocTool:
    """向后兼容：旧代码可能引用 _doc()。"""
    return _tool()  # type: ignore[return-value]


def _wrong_kind_msg(tool_name: str, expected: str, hint: str = "") -> str:
    """当前会话格式与工具不符时的友好提示（而非 AttributeError 崩溃）。"""
    actual = session_doc_kind()
    logger.warning("工具与文档类型不匹配: %s 需要 %s，当前 %s", tool_name, expected, actual)
    hint_str = f" {hint}" if hint else ""
    return (f"{tool_name} 是 {expected.upper()} 专属工具，当前文档是 {actual}。{hint_str}").strip()
