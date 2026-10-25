from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8733
DEFAULT_HOTKEY = "ctrl+alt+o"
DEFAULT_MAX_CELLS = 10_000
DEFAULT_MAX_TEXT_CHARS = 100_000
DEFAULT_REQUEST_TIMEOUT = 30.0
DEFAULT_WRITE_TIMEOUT = 120.0


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


def _env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() not in {"0", "false", "no", "off"}


def default_log_dir() -> Path:
    override = os.environ.get("OFFICE_BRIDGE_LOG_DIR")
    if override:
        return Path(override).expanduser()
    base = os.environ.get("LOCALAPPDATA")
    if base:
        return Path(base) / "office-bridge" / "logs"
    return Path.cwd() / "logs"


@dataclass(slots=True)
class Settings:
    host: str = DEFAULT_HOST
    port: int = DEFAULT_PORT
    hotkey: str = DEFAULT_HOTKEY
    hotkey_enabled: bool = True
    max_cells: int = DEFAULT_MAX_CELLS
    max_text_chars: int = DEFAULT_MAX_TEXT_CHARS
    request_timeout: float = DEFAULT_REQUEST_TIMEOUT
    write_timeout: float = DEFAULT_WRITE_TIMEOUT
    log_level: str = "INFO"
    log_dir: Path = field(default_factory=default_log_dir)
    beep: bool = True
    launch_agent: bool = True
    blocked_calls: str = ""

    @classmethod
    def from_env(
        cls,
        *,
        host: str | None = None,
        port: int | None = None,
        hotkey: str | None = None,
        hotkey_enabled: bool | None = None,
        log_level: str | None = None,
    ) -> "Settings":
        settings = cls(
            host=host or os.environ.get("OFFICE_BRIDGE_HOST", DEFAULT_HOST),
            port=port if port is not None else _env_int("OFFICE_BRIDGE_PORT", DEFAULT_PORT),
            hotkey=hotkey or os.environ.get("OFFICE_BRIDGE_HOTKEY", DEFAULT_HOTKEY),
            hotkey_enabled=(
                hotkey_enabled
                if hotkey_enabled is not None
                else _env_bool("OFFICE_BRIDGE_HOTKEY_ENABLED", True)
            ),
            max_cells=_env_int("OFFICE_BRIDGE_MAX_CELLS", DEFAULT_MAX_CELLS),
            max_text_chars=_env_int("OFFICE_BRIDGE_MAX_TEXT_CHARS", DEFAULT_MAX_TEXT_CHARS),
            request_timeout=_env_float("OFFICE_BRIDGE_REQUEST_TIMEOUT", DEFAULT_REQUEST_TIMEOUT),
            write_timeout=_env_float("OFFICE_BRIDGE_WRITE_TIMEOUT", DEFAULT_WRITE_TIMEOUT),
            log_level=(log_level or os.environ.get("OFFICE_BRIDGE_LOG_LEVEL", "INFO")).upper(),
            beep=_env_bool("OFFICE_BRIDGE_BEEP", True),
            launch_agent=_env_bool("OFFICE_BRIDGE_LAUNCH_AGENT", True),
            blocked_calls=os.environ.get("OFFICE_BRIDGE_BLOCKED_CALLS", ""),
        )
        if settings.host not in {"127.0.0.1", "localhost", "::1"}:
            raise ValueError("The bridge only supports loopback hosts (127.0.0.1, localhost, ::1).")
        return settings

    @property
    def base_url(self) -> str:
        host = "127.0.0.1" if self.host == "localhost" else self.host
        if ":" in host and not host.startswith("["):
            host = f"[{host}]"
        return f"http://{host}:{self.port}"
