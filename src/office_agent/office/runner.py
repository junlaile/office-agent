"""officecli subprocess 底层执行器与基础设施。

被 DocTool / ExcelTool / PptxTool 三个高层工具类共享，是 officecli 封装的
最底层。所有真实 subprocess 调用都收敛到 ``_Runner.run`` 这一个窄口，便于
测试时注入 FakeRunner。

经实测验证的关键点（officecli v1.0.138）:
    - subprocess 必须用 encoding='utf-8'，且 JSON 通过 --commands argv 传递
      （stdin 传中文 JSON 会被误解析；shell echo 也会转码）。
    - 设置 OFFICECLI_NO_AUTO_RESIDENT=1 避免后台常驻进程文件锁。
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
from dataclasses import dataclass, field
from typing import Any

from office_agent.config import settings

logger = logging.getLogger(__name__)


class OfficeCLIError(RuntimeError):
    """officecli 调用失败。"""

    def __init__(
        self,
        message: str,
        *,
        cmd: list[str] | None = None,
        returncode: int | None = None,
        stderr: str | None = None,
    ) -> None:
        super().__init__(message)
        self.cmd = cmd
        self.returncode = returncode
        self.stderr = stderr


def resolve_bin() -> str:
    """按优先级解析 officecli 可执行文件路径。

    1. 环境变量 OFFICECLI_BIN
    2. 工程内 bin/officecli 或 bin/officecli.exe
    3. PATH 中的 officecli
    """
    explicit = settings.officecli_bin.strip()
    if explicit:
        p = os.path.normpath(explicit)
        if os.path.exists(p):
            return p
        found = shutil.which(explicit)
        if found:
            return found
        raise OfficeCLIError(f"OFFICECLI_BIN 指定的路径不存在: {explicit}")

    candidates = [
        settings.project_root / "bin" / "officecli.exe",
        settings.project_root / "bin" / "officecli",
    ]
    for c in candidates:
        if c.exists():
            return str(c)

    found = shutil.which("officecli") or shutil.which("officecli.exe")
    if found:
        return found

    raise OfficeCLIError(
        "找不到 officecli 可执行文件。请执行:\n"
        "    python scripts/fetch_officecli.py\n"
        "下载到工程内 bin/，或在 .env 中设置 OFFICECLI_BIN 指向已有二进制。",
    )


@dataclass
class _Runner:
    """底层 subprocess 执行器。"""

    bin_path: str = field(default_factory=resolve_bin)

    def run(self, args: list[str], *, json_output: bool = False, timeout: int | None = None) -> Any:
        """执行一条 officecli 命令。

        args: 不含可执行文件本身的参数列表，如 ["create", "a.docx"]
        json_output: 是否追加 --json 并解析返回值
        返回: json_output=True 时返回解析后的对象；否则返回 stdout 文本
        """
        cmd = [self.bin_path, *args]
        if json_output and "--json" not in args:
            cmd.append("--json")

        env = dict(os.environ)
        # 关闭常驻模式，避免文件锁/后台进程干扰（每条命令独立 open/save）
        env.setdefault("OFFICECLI_NO_AUTO_RESIDENT", "1")

        try:
            proc = subprocess.run(
                cmd,
                env=env,
                capture_output=True,
                text=True,
                encoding="utf-8",  # 强制 UTF-8，避免 Windows GBK 乱码
                timeout=timeout or settings.officecli_timeout,
                check=False,
            )
        except FileNotFoundError as e:
            logger.debug("officecli 可执行文件不存在: %s", cmd[0])
            raise OfficeCLIError(f"无法执行 officecli: {e}", cmd=cmd) from e
        except subprocess.TimeoutExpired as e:
            logger.debug("officecli 命令超时（%ss）: %s", e.timeout, args[:3])
            raise OfficeCLIError(
                f"officecli 命令超时（{e.timeout}s）: {' '.join(args[:3])}...",
                cmd=cmd,
            ) from e

        if proc.returncode != 0:
            logger.debug(
                "officecli 命令失败 returncode=%s, cmd=%s, stderr=%s",
                proc.returncode,
                args[:4],
                proc.stderr.strip(),
            )
            raise OfficeCLIError(
                f"officecli 返回非零状态 {proc.returncode}: {' '.join(args[:4])}\n"
                f"stderr: {proc.stderr.strip()}",
                cmd=cmd,
                returncode=proc.returncode,
                stderr=proc.stderr.strip(),
            )

        logger.debug("officecli ok: %s", args[:4])

        stdout = proc.stdout
        if json_output:
            stripped = stdout.strip()
            if not stripped:
                return None
            try:
                return json.loads(stripped)
            except json.JSONDecodeError:
                return stdout
        return stdout


# 模块级单例（延迟初始化，避免 import 时就要求二进制存在）
_runner: _Runner | None = None


def get_runner() -> _Runner:
    global _runner
    if _runner is None:
        _runner = _Runner()
    return _runner


def reset_runner(bin_path: str | None = None) -> None:
    """重置 runner，供测试或显式指定路径时使用。"""
    global _runner
    _runner = _Runner(bin_path=bin_path) if bin_path else _Runner()


# ============================================================
# 共享小工具（DocTool / ExcelTool / PptxTool 复用）
# ============================================================
def props_args(props: dict[str, Any]) -> list[str]:
    """把 props 字典展开成 officecli 的 ``--prop k=v`` 参数序列。

    值为 None 的键跳过（表示"不设置该属性"）。
    """
    args: list[str] = []
    for k, v in props.items():
        if v is None:
            continue
        args += ["--prop", f"{k}={v}"]
    return args


def last_child_index(doc_path: str, *, root: str, tag: str) -> int:
    """查询 root 下最后一个 ``tag[N]`` 类子元素的索引（1-based）。无则返回 0。

    统一 DocTool._last_paragraph_index / _last_table_index 与
    PptxTool.last_slide_index 的同构实现：``get <doc> <root> --depth 1``
    后解析 children 路径里的 ``tag[N]`` 取最大 N。

    带 @paraId 等非数字寻址的路径跳过；任何异常按 0 处理（调用方兜底）。
    """
    marker = f"{tag}["
    try:
        data = get_runner().run(
            ["get", doc_path, root, "--depth", "1"],
            json_output=True,
        )
        children = (
            data.get("data", {}).get("results", [{}])[0].get("children", [])
            if isinstance(data, dict)
            else []
        )
        indices = []
        for c in children:
            p = c.get("path", "")
            if marker not in p:
                continue
            seg = p.rsplit(marker, 1)[1].rstrip("]")
            if seg.isdigit():
                indices.append(int(seg))
        return max(indices) if indices else 0
    except Exception:  # noqa: BLE001
        return 0


# ============================================================
# 受限逃生口：白名单 raw
# ============================================================
_RAW_WHITELIST = {
    "create",
    "add",
    "set",
    "get",
    "query",
    "view",
    "close",
    "open",
    "save",
    "validate",
    "batch",
    "merge",
}


def raw(command_args: list[str], *, json_output: bool = False) -> Any:
    """受限逃生口：允许调用白名单内的子命令。

    command_args: 不含可执行文件的完整参数，如 ["view", "a.docx", "outline"]
    第一个元素必须是白名单中的子命令名。
    """
    if not command_args:
        raise OfficeCLIError("raw 调用不能为空")
    sub = command_args[0]
    if sub not in _RAW_WHITELIST:
        raise OfficeCLIError(
            f"raw 子命令 '{sub}' 不在白名单内，允许: {sorted(_RAW_WHITELIST)}",
        )
    return get_runner().run(command_args, json_output=json_output)


def merge_template(template_path: str, output_path: str, data: dict[str, str]) -> str:
    """用 JSON 数据填充模板里的 ``{{key}}`` 占位符，输出到 output_path。

    用于公文模板预填版头固定槽位（发文机关/发文字号/日期等）。
    保留模板段落格式，只替换占位符文本。

    参数:
        template_path: 含 ``{{key}}`` 占位的模板文件路径。
        output_path: 输出文件路径（已存在则覆盖）。
        data: 占位符键值映射，如 ``{"org": "北京市公安局"}``。

    返回 officecli 的 stdout（含 "Replaced keys: N"）。
    """
    payload = json.dumps(data, ensure_ascii=False)
    return get_runner().run(
        [
            "merge",
            template_path,
            output_path,
            "--data",
            payload,
            "--force",  # 覆盖已存在的 output（公文模式下 main.py 已先复制过）
        ]
    )
