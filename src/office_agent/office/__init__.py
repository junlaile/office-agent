"""OfficeCLI 实现层：subprocess runner + Doc/Excel/Pptx 工具类。"""

from .doc import DocTool
from .excel import ExcelTool
from .pptx import PptxTool
from .runner import OfficeCLIError, get_runner, merge_template, reset_runner, resolve_bin

__all__ = [
    "DocTool",
    "ExcelTool",
    "OfficeCLIError",
    "PptxTool",
    "get_runner",
    "merge_template",
    "reset_runner",
    "resolve_bin",
]
