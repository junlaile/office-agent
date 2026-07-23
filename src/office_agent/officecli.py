"""OfficeCLI Python 封装（向后兼容门面）。

历史上本模块包含全部 officecli 封装代码（~1580 行）。现已按职责拆分为:

    - :mod:`office_agent.office.runner`: subprocess 执行器、``OfficeCLIError`` 等。
    - :mod:`office_agent.office.doc`:    Word 文档操作（``DocTool``）。
    - :mod:`office_agent.office.excel`:  Excel 工作簿操作（``ExcelTool``）。
    - :mod:`office_agent.office.pptx`:   PowerPoint 演示文稿操作（``PptxTool``）。

本文件保留为**向后兼容门面**——所有公开符号从这里 re-export，现有代码的
``from office_agent.officecli import X`` 不需要任何改动。

新的代码应直接从对应职责模块 import，例如::

    from office_agent.office.doc import DocTool
    from office_agent.office.runner import OfficeCLIError, merge_template
"""

from __future__ import annotations

# Re-export 所有公开符号，保兼容
from .office.doc import DocTool  # noqa: F401
from .office.excel import ExcelTool  # noqa: F401
from .office.pptx import PptxTool  # noqa: F401
from .office.runner import (  # noqa: F401
    _RAW_WHITELIST,
    OfficeCLIError,
    _Runner,
    get_runner,
    merge_template,
    raw,
    reset_runner,
    resolve_bin,
)

__all__ = [
    "OfficeCLIError",
    "DocTool",
    "ExcelTool",
    "PptxTool",
    "resolve_bin",
    "get_runner",
    "reset_runner",
    "raw",
    "merge_template",
]
