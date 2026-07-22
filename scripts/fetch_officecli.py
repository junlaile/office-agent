"""跨平台下载 OfficeCLI 二进制到工程内 bin/ 目录。

用法:
    python scripts/fetch_officecli.py

说明:
    - 自动识别 OS / 架构，从 GitHub Releases 下载对应 asset
    - Windows 下载后重命名为 officecli.exe；Linux/macOS 重命名为 officecli 并加可执行位
    - 校验 SHA256（下载 SHA256SUMS 比对）
    - 下载失败时打印代理设置提示
    - 如需走代理：set HTTPS_PROXY / HTTP_PROXY 环境变量后重试
"""

from __future__ import annotations

import hashlib
import logging
import os
import platform
import stat
import sys
from pathlib import Path

import httpx

logger = logging.getLogger("fetch_officecli")

REPO = "iOfficeAI/OfficeCLI"
# GitHub API 取 latest release
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"
# 工程根目录（scripts/ 的上一级）
PROJECT_ROOT = Path(__file__).resolve().parent.parent
BIN_DIR = PROJECT_ROOT / "bin"

# OS -> asset 关键字；架构 -> asset 关键字
OS_KEYS = {
    "Windows": "win",
    "Linux": "linux",
    "Darwin": "mac",
}
ARCH_KEYS = {
    "x86_64": "x64",
    "AMD64": "x64",
    "x64": "x64",
    "aarch64": "arm64",
    "arm64": "arm64",
    "ARM64": "arm64",
}


def detect_asset() -> str:
    os_name = platform.system()
    machine = platform.machine()
    if os_name not in OS_KEYS:
        raise SystemExit(f"不支持的操作系统: {os_name}")
    os_key = OS_KEYS[os_name]
    arch_key = ARCH_KEYS.get(machine)
    if arch_key is None:
        raise SystemExit(f"不支持的架构: {machine}")
    return f"officecli-{os_key}-{arch_key}"


def get_latest_release() -> dict:
    headers = {"Accept": "application/vnd.github+json"}
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        headers["Authorization"] = f"Bearer {token}"
    with httpx.Client(timeout=30, follow_redirects=True) as client:
        r = client.get(LATEST_API, headers=headers)
        if r.status_code == 403 and "rate limit" in r.text.lower():
            raise SystemExit(
                "GitHub API 限流。请设置 GITHUB_TOKEN 环境变量后重试，"
                "或稍后再试。"
            )
        r.raise_for_status()
        return r.json()


def pick_asset(release: dict, asset_base: str) -> dict:
    """从 release 中选出匹配的 asset（Windows 带 .exe）。"""
    is_win = platform.system() == "Windows"
    target = f"{asset_base}.exe" if is_win else asset_base
    # alpine 变体特殊处理：非 alpine 系统避免误选 alpine asset
    for a in release.get("assets", []):
        name = a["name"]
        if name == target:
            return a
    # 兜底：包含关键字且不冲突
    for a in release.get("assets", []):
        name = a["name"]
        if name.startswith(asset_base) and "alpine" not in name:
            return a
    raise SystemExit(f"在 release 中找不到匹配的 asset: {target}")


def download(url: str, dest: Path, label: str) -> None:
    logger.info("下载 %s", label)
    logger.info("       <- %s", url)
    logger.info("       -> %s", dest)
    with httpx.Client(timeout=None, follow_redirects=True) as client:
        with client.stream("GET", url) as resp:
            resp.raise_for_status()
            total = int(resp.headers.get("content-length", 0))
            done = 0
            with open(dest, "wb") as f:
                for chunk in resp.iter_bytes(chunk_size=1 << 16):
                    f.write(chunk)
                    done += len(chunk)
                    if total:
                        pct = done * 100 // total
                        bar = "#" * (pct // 2) + "." * (50 - pct // 2)
                        # 进度条只能直接写 stdout（日志会破坏 \r 刷新）
                        sys.stdout.write(f"\r       [{bar}] {pct:3d}% ")
                        sys.stdout.flush()
            sys.stdout.write("\n")
            sys.stdout.flush()
    size_mb = dest.stat().st_size / (1 << 20)
    logger.info("       完成 %.1f MB", size_mb)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_checksum(asset_file: str, binary_path: Path, release: dict) -> bool:
    """下载 SHA256SUMS 比对。失败不阻断（有些 release 可能无此文件）。"""
    sums_asset = next(
        (a for a in release.get("assets", []) if a["name"] == "SHA256SUMS"), None
    )
    if not sums_asset:
        logger.info("release 未提供 SHA256SUMS，跳过校验")
        return True
    try:
        with httpx.Client(timeout=30, follow_redirects=True) as client:
            r = client.get(sums_asset["browser_download_url"])
            r.raise_for_status()
            for line in r.text.splitlines():
                parts = line.split()
                if len(parts) == 2 and parts[1].lstrip("*").endswith(asset_file):
                    expected = parts[0]
                    actual = sha256(binary_path)
                    if actual != expected:
                        logger.warning("SHA256 校验失败: 期望 %s, 实际 %s", expected, actual)
                        return False
                    logger.info("SHA256 校验通过")
                    return True
        logger.info("SHA256SUMS 中未找到对应条目，跳过")
        return True
    except Exception as e:  # noqa: BLE001
        logger.warning("SHA256 校验跳过（%s）", e)
        return True


def main() -> int:
    # 此脚本独立运行，自配日志（不依赖 office_agent.config）
    logging.basicConfig(
        level=os.environ.get("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    BIN_DIR.mkdir(parents=True, exist_ok=True)
    asset_base = detect_asset()
    logger.info("检测平台: %s %s -> asset 基名: %s",
                platform.system(), platform.machine(), asset_base)

    try:
        release = get_latest_release()
    except Exception as e:  # noqa: BLE001
        logger.error("获取最新 release 失败: %s", e)
        proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
        if not proxy:
            logger.info(
                "若网络受限，请设置代理后重试：\n"
                "       Windows (PowerShell): $env:HTTPS_PROXY='http://127.0.0.1:7890'\n"
                "       Linux/macOS: export HTTPS_PROXY=http://127.0.0.1:7890"
            )
        return 1

    tag = release.get("tag_name", "unknown")
    logger.info("最新 release: %s", tag)

    asset = pick_asset(release, asset_base)
    asset_name = asset["name"]

    is_win = platform.system() == "Windows"
    final_name = "officecli.exe" if is_win else "officecli"
    final_path = BIN_DIR / final_name

    download(asset["browser_download_url"], final_path, asset_name)
    verify_checksum(asset_name, final_path, release)

    if not is_win:
        final_path.chmod(final_path.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    logger.info("完成: %s", final_path)
    logger.info("       运行 '%s --version' 验证。", final_path)
    logger.info("       OFFICECLI_BIN 可指向该路径，或将其加入 PATH。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
