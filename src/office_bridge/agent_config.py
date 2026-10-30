from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Sequence

DEFAULT_AGENT = "opencode"
AGENT_ENV = "OFFICE_BRIDGE_AGENT"
CONFIG_ENV = "OFFICE_BRIDGE_AGENT_CONFIG"
SKILL_SOURCE_ENV = "OFFICE_BRIDGE_SKILL_SOURCE"
SKILL_NAME = "office"


class AgentConfigError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class AgentPreset:
    name: str
    label: str
    command: tuple[str, ...]
    project_skill: Path | None
    user_skill: Path | None
    default_scope: str
    preload_args: tuple[str, ...] = ()


@dataclass(slots=True)
class AgentConfig:
    agent: str = DEFAULT_AGENT
    command: list[str] | None = None
    args: list[str] = field(default_factory=list)
    skill_scope: str | None = None
    skill_dir: str | None = None
    install_skill: bool = True
    preload_skill: bool = True

    def to_dict(self) -> dict[str, Any]:
        return {
            "agent": self.agent,
            "command": self.command,
            "args": self.args,
            "skill_scope": self.skill_scope,
            "skill_dir": self.skill_dir,
            "install_skill": self.install_skill,
            "preload_skill": self.preload_skill,
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "AgentConfig":
        if not isinstance(payload, dict):
            raise AgentConfigError("The agent configuration must be a JSON object.")

        agent = payload.get("agent", DEFAULT_AGENT)
        if not isinstance(agent, str) or not agent.strip():
            raise AgentConfigError("'agent' must be a non-empty string.")

        command = payload.get("command")
        if command is not None:
            command = _string_list(command, "command")
            if not command:
                command = None

        args = _string_list(payload.get("args", []), "args")
        skill_scope = payload.get("skill_scope")
        if skill_scope not in {None, "project", "user"}:
            raise AgentConfigError("'skill_scope' must be 'project', 'user', or null.")

        skill_dir = payload.get("skill_dir")
        if skill_dir is not None and not isinstance(skill_dir, str):
            raise AgentConfigError("'skill_dir' must be a string or null.")

        return cls(
            agent=agent.strip(),
            command=command,
            args=args,
            skill_scope=skill_scope,
            skill_dir=skill_dir,
            install_skill=_as_bool(payload.get("install_skill", True), "install_skill"),
            preload_skill=_as_bool(payload.get("preload_skill", True), "preload_skill"),
        )


def _string_list(value: Any, field_name: str) -> list[str]:
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise AgentConfigError(f"'{field_name}' must be a list of strings.")
    return list(value)


def _as_bool(value: Any, field_name: str) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() not in {"0", "false", "no", "off"}
    raise AgentConfigError(f"'{field_name}' must be a boolean.")


def _home() -> Path:
    return Path.home()


def _hermes_home() -> Path:
    override = os.environ.get("HERMES_HOME")
    if override:
        return Path(override).expanduser()
    default = _home() / ".hermes"
    if default.exists():
        return default
    local = os.environ.get("LOCALAPPDATA")
    if local and (Path(local) / "hermes").exists():
        return Path(local) / "hermes"
    return default


def _opencode_config_home() -> Path:
    override = os.environ.get("XDG_CONFIG_HOME")
    base = Path(override).expanduser() if override else _home() / ".config"
    return base / "opencode"


def agent_presets() -> dict[str, AgentPreset]:
    return {
        "opencode": AgentPreset(
            name="opencode",
            label="OpenCode",
            command=("opencode",),
            project_skill=Path(".opencode") / "skills" / SKILL_NAME,
            user_skill=_opencode_config_home() / "skills" / SKILL_NAME,
            default_scope="project",
        ),
        "claude": AgentPreset(
            name="claude",
            label="Claude Code",
            command=("claude",),
            project_skill=Path(".claude") / "skills" / SKILL_NAME,
            user_skill=_home() / ".claude" / "skills" / SKILL_NAME,
            default_scope="project",
        ),
        "codex": AgentPreset(
            name="codex",
            label="Codex",
            command=("codex",),
            project_skill=Path(".agents") / "skills" / SKILL_NAME,
            user_skill=_home() / ".agents" / "skills" / SKILL_NAME,
            default_scope="project",
        ),
        "hermes": AgentPreset(
            name="hermes",
            label="Hermes Agent",
            command=("hermes",),
            project_skill=Path(".hermes") / "skills" / SKILL_NAME,
            user_skill=_hermes_home() / "skills" / SKILL_NAME,
            default_scope="user",
            preload_args=("-s", SKILL_NAME),
        ),
        "custom": AgentPreset(
            name="custom",
            label="Custom command",
            command=(),
            project_skill=None,
            user_skill=None,
            default_scope="project",
        ),
    }


def default_config_path() -> Path:
    override = os.environ.get(CONFIG_ENV)
    if override:
        return Path(override).expanduser()
    base = os.environ.get("LOCALAPPDATA")
    if base:
        return Path(base) / "office-bridge" / "agent.json"
    return _home() / ".office-bridge" / "agent.json"


def load_agent_config(path: str | Path | None = None) -> AgentConfig:
    config_path = Path(path).expanduser() if path else default_config_path()
    if config_path.exists():
        try:
            payload = json.loads(config_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise AgentConfigError(f"Could not read {config_path}: {exc}") from exc
        config = AgentConfig.from_dict(payload)
    else:
        config = AgentConfig()

    override = os.environ.get(AGENT_ENV)
    if override and override.strip() != config.agent:
        config = replace(
            config,
            agent=override.strip(),
            command=None,
            skill_scope=None,
            skill_dir=None,
            install_skill=True,
            preload_skill=True,
        )
    get_preset(config)
    return config


def save_agent_config(config: AgentConfig, path: str | Path | None = None) -> Path:
    get_preset(config)
    config_path = Path(path).expanduser() if path else default_config_path()
    config_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = config_path.with_name(config_path.name + ".tmp")
    temporary.write_text(
        json.dumps(config.to_dict(), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, config_path)
    return config_path


def get_preset(config: AgentConfig) -> AgentPreset:
    presets = agent_presets()
    preset = presets.get(config.agent)
    if preset is None:
        available = ", ".join(sorted(presets))
        raise AgentConfigError(
            f"Unknown agent '{config.agent}'. Choose one of: {available}."
        )
    return preset


def resolve_command(config: AgentConfig) -> list[str]:
    preset = get_preset(config)
    command = list(config.command) if config.command else list(preset.command)
    if not command:
        raise AgentConfigError(
            f"No launch command is configured for agent '{config.agent}'."
        )
    return command


def resolve_skill_dir(
    config: AgentConfig,
    project_root: str | Path | None = None,
) -> Path:
    preset = get_preset(config)
    root = Path(project_root or Path.cwd()).expanduser().resolve()
    if config.skill_dir:
        path = Path(config.skill_dir).expanduser()
        return path if path.is_absolute() else (root / path).resolve()

    scope = config.skill_scope or preset.default_scope
    if scope == "project":
        if preset.project_skill is None:
            raise AgentConfigError(
                f"Agent '{config.agent}' has no project skill directory. "
                "Set --skill-dir or choose user scope."
            )
        return (root / preset.project_skill).resolve()
    if scope == "user":
        if preset.user_skill is None:
            raise AgentConfigError(
                f"Agent '{config.agent}' has no user skill directory. "
                "Set --skill-dir."
            )
        return preset.user_skill.expanduser()
    raise AgentConfigError("'skill_scope' must be 'project' or 'user'.")


def skill_source_path(source: str | Path | None = None) -> Path:
    override = source or os.environ.get(SKILL_SOURCE_ENV)
    if override:
        path = Path(override).expanduser()
        return path if path.is_absolute() else (Path.cwd() / path).resolve()

    candidates = [
        Path(__file__).resolve().parents[2] / "skills" / SKILL_NAME / "SKILL.md",
        Path.cwd() / "skills" / SKILL_NAME / "SKILL.md",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise AgentConfigError(
        "Could not locate skills/office/SKILL.md. "
        f"Set {SKILL_SOURCE_ENV} to the source file."
    )


def install_skill(
    config: AgentConfig,
    project_root: str | Path | None = None,
    source: str | Path | None = None,
) -> Path | None:
    if not config.install_skill:
        return None
    preset = get_preset(config)
    if not config.skill_dir and preset.project_skill is None and preset.user_skill is None:
        return None
    source_path = skill_source_path(source)
    if not source_path.is_file():
        raise AgentConfigError(f"Skill source not found: {source_path}")
    destination = resolve_skill_dir(config, project_root)
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / "SKILL.md"
    if source_path.resolve() != target.resolve():
        shutil.copy2(source_path, target)
    return target


def _wrap_executable(path: str) -> list[str]:
    suffix = Path(path).suffix.lower()
    if suffix == ".ps1":
        return [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            path,
        ]
    if suffix in {".cmd", ".bat"}:
        return [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/c", path]
    return [path]


def _resolve_executable(name: str) -> str:
    candidate = Path(name).expanduser()
    if candidate.is_file():
        return str(candidate.resolve())
    found = shutil.which(name)
    if found:
        return found
    if os.name == "nt" and not candidate.suffix:
        for extension in (".exe", ".cmd", ".bat", ".ps1"):
            found = shutil.which(name + extension)
            if found:
                return found
    raise AgentConfigError(f"Executable '{name}' was not found on PATH.")


def build_launch_argv(
    config: AgentConfig,
    *,
    resolve: bool = True,
) -> list[str]:
    preset = get_preset(config)
    command = resolve_command(config)
    executable, extra = command[0], command[1:]
    args = [*extra, *config.args]
    if config.preload_skill:
        args.extend(preset.preload_args)
    if not resolve:
        return [executable, *args]
    return [*_wrap_executable(_resolve_executable(executable)), *args]


def describe_agent(
    config: AgentConfig,
    project_root: str | Path | None = None,
) -> dict[str, Any]:
    preset = get_preset(config)
    data: dict[str, Any] = {
        "ok": False,
        "agent": config.agent,
        "display_name": preset.label,
        "config_path": str(default_config_path()),
        "command": None,
        "launch_argv": None,
        "skill_dir": None,
        "skill_installed": False,
        "error": None,
    }
    try:
        data["command"] = build_launch_argv(config, resolve=False)
    except AgentConfigError as exc:
        data["error"] = str(exc)
        return data

    try:
        data["launch_argv"] = build_launch_argv(config, resolve=True)
    except AgentConfigError as exc:
        data["error"] = str(exc)
        return data

    has_skill_location = (
        config.skill_dir is not None
        or preset.project_skill is not None
        or preset.user_skill is not None
    )
    if config.install_skill and has_skill_location:
        try:
            skill_dir = resolve_skill_dir(config, project_root)
            data["skill_dir"] = str(skill_dir)
            data["skill_installed"] = (skill_dir / "SKILL.md").is_file()
        except AgentConfigError as exc:
            data["error"] = str(exc)
            return data
    data["ok"] = True
    return data


def _read_config_from_args(args: argparse.Namespace) -> AgentConfig:
    config = load_agent_config()
    if getattr(args, "agent", None):
        config = replace(
            config,
            agent=args.agent,
            command=None,
            skill_scope=None,
            skill_dir=None,
            install_skill=True,
            preload_skill=True,
        )
    if getattr(args, "scope", None):
        config = replace(config, skill_scope=args.scope)
    if getattr(args, "skill_dir", None):
        config = replace(config, skill_dir=args.skill_dir)
    return config


def _cmd_list(args: argparse.Namespace) -> int:
    config = load_agent_config()
    presets = agent_presets()
    print(f"Configuration: {default_config_path()}")
    print(f"Current agent: {config.agent}")
    print()
    for name, preset in presets.items():
        marker = "*" if name == config.agent else " "
        command = " ".join(preset.command) or "(set --command)"
        print(
            f"{marker} {name:<8} {preset.label:<14} "
            f"command={command!r} skill={preset.default_scope}"
        )
    return 0


def _cmd_show(args: argparse.Namespace) -> int:
    config = load_agent_config()
    data = describe_agent(config)
    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0 if data["ok"] else 1
    print(f"Config:   {data['config_path']}")
    print(f"Agent:    {data['agent']} ({data['display_name']})")
    print(f"Command:  {data['command']}")
    print(f"Launch:   {data['launch_argv']}")
    print(f"Skill:    {data['skill_dir'] or '(disabled)'}")
    print(f"Installed:{' yes' if data['skill_installed'] else ' no'}")
    if data["error"]:
        print(f"Error:    {data['error']}")
    return 0 if data["ok"] else 1


def _cmd_set(args: argparse.Namespace) -> int:
    presets = agent_presets()
    if args.agent not in presets:
        available = ", ".join(sorted(presets))
        raise AgentConfigError(
            f"Unknown agent '{args.agent}'. Choose one of: {available}."
        )
    install_skill_enabled = (
        args.install_skill
        if args.install_skill is not None
        else args.agent != "custom"
    )
    config = AgentConfig(
        agent=args.agent,
        command=[args.command] if args.command else None,
        args=list(args.arg or []),
        skill_scope=args.scope,
        skill_dir=args.skill_dir,
        install_skill=install_skill_enabled,
        preload_skill=not args.no_preload,
    )
    path = save_agent_config(config)
    print(f"Saved {path}")
    print(f"Agent: {config.agent}")
    if config.command:
        print(f"Command: {' '.join(config.command)}")
    if args.install:
        installed = install_skill(config)
        if installed:
            print(f"Installed skill: {installed}")
        else:
            print("Skill installation is disabled for this agent.")
    return 0


def _cmd_install_skill(args: argparse.Namespace) -> int:
    config = _read_config_from_args(args)
    installed = install_skill(
        config,
        project_root=args.project_root,
        source=args.source,
    )
    if installed is None:
        print("Skill installation is disabled for this agent.")
        return 0
    print(installed)
    return 0


def _cmd_resolve(args: argparse.Namespace) -> int:
    config = load_agent_config()
    data = describe_agent(config, project_root=args.project_root)
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


def _cmd_reset(args: argparse.Namespace) -> int:
    path = default_config_path()
    if path.exists():
        path.unlink()
        print(f"Removed {path}")
    else:
        print(f"No configuration at {path}")
    return 0


def _cmd_setup(args: argparse.Namespace) -> int:
    config = load_agent_config()
    presets = agent_presets()
    print("Office bridge agent setup")
    print(f"Current agent: {config.agent}")
    print("Available agents: " + ", ".join(presets))
    try:
        choice = input(f"Agent [{config.agent}]: ").strip() or config.agent
        if choice not in presets:
            raise AgentConfigError(
                f"Unknown agent '{choice}'. Choose one of: {', '.join(presets)}."
            )
        preset = presets[choice]
        command = None
        if choice == "custom":
            command = input("Command (for example C:\\Tools\\agent.exe): ").strip()
            if not command:
                raise AgentConfigError("A command is required for the custom agent.")
        scope = (
            input(
                f"Skill scope (project/user) [{preset.default_scope}]: "
            ).strip()
            or preset.default_scope
        )
        if scope not in {"project", "user"}:
            raise AgentConfigError("Skill scope must be 'project' or 'user'.")
        config = AgentConfig(
            agent=choice,
            command=[command] if command else None,
            skill_scope=scope,
            install_skill=choice != "custom",
        )
        path = save_agent_config(config)
        print(f"Saved {path}")
        if config.install_skill:
            answer = input("Install the Office skill now? [Y/n]: ").strip().lower()
            if answer in {"", "y", "yes"}:
                installed = install_skill(config)
                if installed:
                    print(f"Installed skill: {installed}")
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        return 1
    print("Start the configured agent with start-agent.bat.")
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="office-agent",
        description="Configure which AI agent the Office bridge launches.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    list_parser = sub.add_parser("list", help="List available agents.")
    list_parser.set_defaults(handler=_cmd_list)

    show_parser = sub.add_parser("show", help="Show the current agent configuration.")
    show_parser.add_argument("--json", action="store_true", help="Emit JSON.")
    show_parser.set_defaults(handler=_cmd_show)

    set_parser = sub.add_parser("set", help="Select an agent.")
    set_parser.add_argument("agent", help="opencode, claude, codex, hermes, or custom")
    set_parser.add_argument(
        "--command",
        help="Executable path or name for a custom agent.",
    )
    set_parser.add_argument(
        "--arg",
        action="append",
        default=[],
        help=(
            "Extra launch argument. Repeat for multiple arguments. "
            "Use --arg=--flag when the value starts with a dash."
        ),
    )
    set_parser.add_argument(
        "--scope",
        choices=["project", "user"],
        help="Where to install the Office skill.",
    )
    set_parser.add_argument("--skill-dir", help="Explicit skill destination directory.")
    set_parser.add_argument(
        "--install-skill",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Install the Office skill when starting the agent.",
    )
    set_parser.add_argument(
        "--no-preload",
        action="store_true",
        help="Do not pass the agent-specific skill preload arguments.",
    )
    set_parser.add_argument(
        "--install",
        action="store_true",
        help="Install the skill immediately after saving.",
    )
    set_parser.set_defaults(handler=_cmd_set)

    install_parser = sub.add_parser(
        "install-skill",
        help="Copy the Office skill to the configured destination.",
    )
    install_parser.add_argument("--agent", help="Override the configured agent.")
    install_parser.add_argument("--scope", choices=["project", "user"])
    install_parser.add_argument("--skill-dir")
    install_parser.add_argument("--project-root", default=str(Path.cwd()))
    install_parser.add_argument("--source", help="Path to the source SKILL.md.")
    install_parser.set_defaults(handler=_cmd_install_skill)

    resolve_parser = sub.add_parser(
        "resolve",
        help="Resolve the launch command and skill path as JSON.",
    )
    resolve_parser.add_argument("--project-root", default=str(Path.cwd()))
    resolve_parser.add_argument(
        "--json",
        action="store_true",
        help="Emit JSON (the default; accepted for scripting symmetry).",
    )
    resolve_parser.set_defaults(handler=_cmd_resolve)

    reset_parser = sub.add_parser("reset", help="Remove the saved configuration.")
    reset_parser.set_defaults(handler=_cmd_reset)

    setup_parser = sub.add_parser("setup", help="Interactively choose an agent.")
    setup_parser.set_defaults(handler=_cmd_setup)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(argv) if argv is not None else sys.argv[1:]
    if not arguments:
        arguments = ["setup"]
    parser = _build_parser()
    args = parser.parse_args(arguments)
    try:
        return int(args.handler(args))
    except (AgentConfigError, OSError) as exc:
        print(f"office-agent: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
