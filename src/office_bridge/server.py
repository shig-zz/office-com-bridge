from __future__ import annotations

import argparse
import ctypes
import logging
import os
import shutil
import socket
import subprocess
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

import uvicorn

from . import __version__
from .api import create_app
from .com_worker import BridgeCore
from .config import Settings
from .errors import BridgeError
from .hotkey import HotkeyError, HotkeyListener

logger = logging.getLogger("office_bridge")

_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_kernel32.CreateMutexW.restype = ctypes.c_void_p
_kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
_kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
_ERROR_ALREADY_EXISTS = 183


def configure_logging(settings: Settings) -> None:
    level = getattr(logging, settings.log_level.upper(), logging.INFO)
    root = logging.getLogger()
    root.setLevel(level)
    for handler in list(root.handlers):
        root.removeHandler(handler)

    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler(stream=sys.stderr)
    console.setFormatter(formatter)
    root.addHandler(console)

    try:
        settings.log_dir.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            settings.log_dir / "bridge.log",
            maxBytes=2_000_000,
            backupCount=3,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)
    except OSError as exc:
        logger.warning("could not open log file in %s: %s", settings.log_dir, exc)

    logging.getLogger("uvicorn.error").setLevel(level)
    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)


def acquire_single_instance(port: int) -> int:
    name = f"Local\\OfficeComBridge-{port}"
    handle = _kernel32.CreateMutexW(None, False, name)
    if not handle:
        raise OSError("Could not create the bridge mutex.")
    if ctypes.get_last_error() == _ERROR_ALREADY_EXISTS:
        _kernel32.CloseHandle(handle)
        raise RuntimeError(
            f"Another Office bridge is already listening for this port ({port})."
        )
    return int(handle)


def _beep(kind: int) -> None:
    try:
        ctypes.WinDLL("user32").MessageBeep(kind)
    except Exception:
        pass


def _find_opencode() -> str | None:
    for name in ("opencode", "opencode.exe", "opencode.cmd", "opencode.ps1"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _launch_agent(cwd: str, url: str) -> None:
    """Start OpenCode in a new terminal window (Windows Terminal when available)."""
    launcher = _find_opencode()
    if launcher is None:
        raise FileNotFoundError("opencode was not found on PATH")
    if launcher.lower().endswith(".ps1"):
        command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", launcher]
    else:
        command = [launcher]
    env = dict(os.environ)
    env["OFFICE_BRIDGE_URL"] = url
    scripts_dir = os.path.dirname(sys.executable)
    env["PATH"] = scripts_dir + os.pathsep + env.get("PATH", "")
    wt = shutil.which("wt")
    if wt:
        subprocess.Popen([wt, "-d", cwd, *command], env=env)
    else:
        subprocess.Popen(
            command,
            cwd=cwd,
            env=env,
            creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
        )


def _port_available(host: str, port: int) -> bool:
    if host in {"localhost", "::1"}:
        family = socket.AF_INET6 if host == "::1" else socket.AF_INET
        bind_host = "::1" if host == "::1" else "127.0.0.1"
    else:
        family = socket.AF_INET
        bind_host = host
    sock = socket.socket(family, socket.SOCK_STREAM)
    try:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.bind((bind_host, port))
        return True
    except OSError:
        return False
    finally:
        sock.close()


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="office-bridge",
        description="Local HTTP bridge to running Microsoft Office applications.",
    )
    parser.add_argument("--host", help="Loopback host (default: 127.0.0.1).")
    parser.add_argument("--port", type=int, help="Port to listen on (default: 8733).")
    parser.add_argument("--hotkey", help="Global hotkey, for example ctrl+alt+o.")
    parser.add_argument(
        "--no-hotkey",
        action="store_true",
        help="Do not register the global hotkey (useful for diagnostics).",
    )
    parser.add_argument("--log-level", help="DEBUG, INFO, WARNING, or ERROR.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        settings = Settings.from_env(
            host=args.host,
            port=args.port,
            hotkey=args.hotkey,
            hotkey_enabled=False if args.no_hotkey else None,
            log_level=args.log_level,
        )
    except ValueError as exc:
        print(f"office-bridge: {exc}", file=sys.stderr)
        return 2

    configure_logging(settings)
    logger.info("office bridge %s starting on %s", __version__, settings.base_url)

    mutex_handle: int | None = None
    try:
        mutex_handle = acquire_single_instance(settings.port)
    except (RuntimeError, OSError) as exc:
        logger.error("%s", exc)
        return 3

    if not _port_available(settings.host, settings.port):
        logger.error(
            "Port %s is already in use by another process. Choose a different port with "
            "OFFICE_BRIDGE_PORT or --port.",
            settings.port,
        )
        _kernel32.CloseHandle(ctypes.c_void_p(mutex_handle))
        return 3

    core = BridgeCore(settings)
    hotkey: HotkeyListener | None = None
    exit_code = 0
    try:
        core.start()

        if settings.hotkey_enabled:
            def on_hotkey() -> None:
                try:
                    result = core.select_foreground()
                    target = result.get("target", {})
                    logger.info(
                        "hotkey target: %s/%s",
                        target.get("application"),
                        target.get("document"),
                    )
                    if settings.launch_agent:
                        try:
                            _launch_agent(str(Path.cwd()), settings.base_url)
                            logger.info("launched OpenCode agent in a new terminal")
                        except Exception:
                            logger.exception("failed to launch the OpenCode agent")
                except BridgeError as exc:
                    logger.warning("hotkey rejected: %s: %s", exc.code, exc.message)
                    if settings.beep:
                        _beep(0x00000030)
                except Exception:
                    logger.exception("hotkey target selection failed")
                    if settings.beep:
                        _beep(0x00000010)

            try:
                hotkey = HotkeyListener(settings.hotkey, on_hotkey)
                hotkey.start()
            except HotkeyError as exc:
                logger.error("could not register the hotkey: %s", exc)
                if hotkey is not None:
                    hotkey.stop()
                return 4
        else:
            logger.warning("hotkey is disabled; no Office target can be selected")

        app = create_app(core, settings)
        logger.info(
            "target selection hotkey: %s",
            settings.hotkey if settings.hotkey_enabled else "(disabled)",
        )
        uvicorn.run(
            app,
            host=settings.host,
            port=settings.port,
            log_config=None,
            access_log=False,
            log_level=settings.log_level.lower(),
        )
    except KeyboardInterrupt:
        logger.info("interrupted")
    except SystemExit as exc:
        code = exc.code
        exit_code = int(code) if isinstance(code, int) else 1
    finally:
        if hotkey is not None:
            hotkey.stop()
        try:
            core.stop()
        except Exception:
            logger.exception("failed to stop the COM worker cleanly")
        if mutex_handle is not None:
            _kernel32.CloseHandle(ctypes.c_void_p(mutex_handle))
        logger.info("office bridge stopped")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
