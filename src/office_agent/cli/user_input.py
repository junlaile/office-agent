"""忙时用户输入桥：单一 stdin 读线程，区分软补充 / 强制打断 / 继续 / 退出。

交互约定:
  - 普通文字 → SUPPLEMENT（排队，下一轮 agent 注入）
  - !内容 / /force 内容 / 强制:内容 → FORCE（节点边界打断后注入）
  - 继续 / continue / 请继续完成 … → CONTINUE
  - 退出 / quit / exit → QUIT

设计要点: 全进程只有读线程调用 input()；主线程通过
blocking_readline() / 忙时自动分类投递，避免与 ask_user 抢 stdin。
"""

from __future__ import annotations

import re
import threading
from collections import deque
from contextvars import ContextVar
from dataclasses import dataclass
from enum import StrEnum
from queue import Empty, Queue


class InputKind(StrEnum):
    CONTINUE = "continue"
    FORCE = "force"
    SUPPLEMENT = "supplement"
    QUIT = "quit"


@dataclass(frozen=True)
class ClassifiedInput:
    kind: InputKind
    text: str = ""


_CONTINUE_RE = re.compile(
    r"^(继续|请继续|请继续完成|请继续完成当前任务|接着|接着做|接着写|"
    r"continue|go\s*on|keep\s*going)[.!！。…]*$",
    re.IGNORECASE,
)

_FORCE_RE = re.compile(
    r"^(?:!|/force\b|强制[:：]?\s*)(.+)$",
    re.IGNORECASE | re.DOTALL,
)

_QUIT_RE = re.compile(
    r"^(退出|结束|quit|exit|q)[.!！。]*$",
    re.IGNORECASE,
)

CONTINUE_PROMPT = "请继续完成当前任务。请基于已有文档内容继续，不要重新 create_doc。"

PREFIX_SUPPLEMENT = "【用户补充】"
PREFIX_FORCE = "【用户强制打断】"


def classify(line: str) -> ClassifiedInput | None:
    """分类一行用户输入。空行返回 None。"""
    text = (line or "").strip()
    if not text:
        return None
    if _QUIT_RE.match(text):
        return ClassifiedInput(InputKind.QUIT)
    if _CONTINUE_RE.match(text):
        return ClassifiedInput(InputKind.CONTINUE, CONTINUE_PROMPT)
    m = _FORCE_RE.match(text)
    if m:
        body = m.group(1).strip()
        if body:
            return ClassifiedInput(InputKind.FORCE, body)
        return None
    return ClassifiedInput(InputKind.SUPPLEMENT, text)


class UserInputBridge:
    """单一 stdin 读线程 + 忙时分类队列。"""

    def __init__(self) -> None:
        self._raw: Queue[str | None] = Queue()  # None = EOF
        self._soft: deque[str] = deque()
        self._force_text: str | None = None
        self._force_event = threading.Event()
        self._continue_pending = False
        self._quit_pending = False
        self._soft_pause = False
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        # busy=True：读到的行自动 classify 进 soft/force
        # busy=False：行进 _raw，由 blocking_readline 取走
        self._busy = False
        self._eof = False

    # ── 生命周期 ──────────────────────────────────────────

    def start(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(
            target=self._reader_loop,
            name="user-input-bridge",
            daemon=True,
        )
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()

    def set_busy(self, busy: bool) -> None:
        with self._lock:
            self._busy = busy

    # ── 主线程阻塞读一行（ask_user / 软暂停 REPL） ────────

    def blocking_readline(self, prompt: str = "") -> str:
        """打印 prompt，阻塞等到下一行。EOF → 抛 EOFError。"""
        if prompt:
            # 与原先 input(prompt) 行为一致：prompt 不换行
            print(prompt, end="", flush=True)
        while True:
            # 先消化忙时残留：若刚从 busy 切回，队列里可能已有行
            try:
                line = self._raw.get(timeout=0.1)
            except Empty:
                if self._eof:
                    raise EOFError from None
                if self._stop.is_set():
                    raise EOFError from None
                continue
            if line is None:
                self._eof = True
                raise EOFError
            return line

    # ── 分类投递 ──────────────────────────────────────────

    def submit(self, line: str) -> ClassifiedInput | None:
        item = classify(line)
        if item is None:
            return None
        with self._lock:
            if item.kind == InputKind.QUIT:
                self._quit_pending = True
            elif item.kind == InputKind.CONTINUE:
                self._continue_pending = True
            elif item.kind == InputKind.FORCE:
                self._force_text = item.text
                self._force_event.set()
            elif item.kind == InputKind.SUPPLEMENT:
                self._soft.append(item.text)
        return item

    def request_soft_pause(self) -> None:
        with self._lock:
            self._soft_pause = True

    # ── 查询 / 消费 ──────────────────────────────────────

    def has_force(self) -> bool:
        return self._force_event.is_set()

    def has_pending(self) -> bool:
        with self._lock:
            return (
                self._force_event.is_set()
                or self._continue_pending
                or self._quit_pending
                or self._soft_pause
                or bool(self._soft)
            )

    def consume_force(self) -> str | None:
        with self._lock:
            if not self._force_event.is_set():
                return None
            text = self._force_text
            self._force_text = None
            self._force_event.clear()
            return text

    def drain_soft(self) -> list[str]:
        with self._lock:
            items = list(self._soft)
            self._soft.clear()
            return items

    def consume_continue(self) -> bool:
        with self._lock:
            if not self._continue_pending:
                return False
            self._continue_pending = False
            return True

    def consume_quit(self) -> bool:
        with self._lock:
            if not self._quit_pending:
                return False
            self._quit_pending = False
            return True

    def consume_soft_pause(self) -> bool:
        with self._lock:
            if not self._soft_pause:
                return False
            self._soft_pause = False
            return True

    def peek_soft_pause(self) -> bool:
        with self._lock:
            return self._soft_pause

    # ── 后台读循环（唯一 input() 调用点） ────────────────

    def _reader_loop(self) -> None:
        while not self._stop.is_set():
            try:
                line = input()
            except EOFError:
                self._eof = True
                self._raw.put(None)
                break
            except KeyboardInterrupt:
                self.request_soft_pause()
                continue
            except Exception:  # noqa: BLE001
                continue

            if self._stop.is_set():
                break

            with self._lock:
                busy = self._busy

            if busy:
                # 忙时：分类进 soft/force/continue/quit
                self.submit(line)
            else:
                # 空闲：交给 blocking_readline
                self._raw.put(line)


_bridge_var: ContextVar[UserInputBridge | None] = ContextVar(
    "user_input_bridge", default=None
)
_default_bridge: UserInputBridge | None = None


def set_bridge(bridge: UserInputBridge | None) -> None:
    global _default_bridge
    _default_bridge = bridge
    _bridge_var.set(bridge)


def get_bridge() -> UserInputBridge | None:
    return _bridge_var.get() or _default_bridge
