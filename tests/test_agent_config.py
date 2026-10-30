from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest

import office_bridge.agent_config as agent_config


def write_skill(tmp_path: Path) -> Path:
    source = tmp_path / "skills" / "office" / "SKILL.md"
    source.parent.mkdir(parents=True)
    source.write_text("---\nname: office\ndescription: test\n---\n", encoding="utf-8")
    return source


def test_default_config_is_opencode(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(tmp_path / "agent.json"))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)

    config = agent_config.load_agent_config()

    assert config.agent == "opencode"
    assert agent_config.build_launch_argv(config, resolve=False) == ["opencode"]
    assert agent_config.resolve_skill_dir(config, tmp_path) == (
        tmp_path / ".opencode" / "skills" / "office"
    ).resolve()


def test_save_and_load_hermes_preset(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(tmp_path / "agent.json"))
    monkeypatch.setenv("HERMES_HOME", str(tmp_path / "hermes"))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)

    path = agent_config.save_agent_config(agent_config.AgentConfig(agent="hermes"))
    loaded = agent_config.load_agent_config()

    assert path == tmp_path / "agent.json"
    assert loaded.agent == "hermes"
    assert agent_config.build_launch_argv(loaded, resolve=False) == [
        "hermes",
        "-s",
        "office",
    ]
    assert agent_config.resolve_skill_dir(loaded, tmp_path) == (
        tmp_path / "hermes" / "skills" / "office"
    )


def test_env_agent_overrides_saved_command(monkeypatch, tmp_path) -> None:
    path = tmp_path / "agent.json"
    agent_config.save_agent_config(
        agent_config.AgentConfig(
            agent="custom",
            command=["C:\\Tools\\agent.exe"],
            install_skill=False,
        ),
        path,
    )
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(path))
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT", "codex")

    loaded = agent_config.load_agent_config()

    assert loaded.agent == "codex"
    assert loaded.command is None
    assert agent_config.build_launch_argv(loaded, resolve=False) == ["codex"]


def test_install_skill_copies_to_project(monkeypatch, tmp_path) -> None:
    source = write_skill(tmp_path)
    config = agent_config.AgentConfig(agent="claude")

    target = agent_config.install_skill(config, tmp_path, source)

    assert target is not None
    assert target == tmp_path / ".claude" / "skills" / "office" / "SKILL.md"
    assert target.read_text(encoding="utf-8") == source.read_text(encoding="utf-8")


def test_install_skill_not_applicable_to_custom_without_dir(monkeypatch, tmp_path) -> None:
    source = write_skill(tmp_path)
    config = agent_config.AgentConfig(
        agent="custom",
        command=[sys.executable, "-V"],
        install_skill=True,
        preload_skill=False,
    )

    assert agent_config.install_skill(config, tmp_path, source) is None

    data = agent_config.describe_agent(config, tmp_path)
    assert data["ok"] is True
    assert data["skill_dir"] is None


def test_custom_agent_requires_command(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(tmp_path / "agent.json"))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)
    config = agent_config.AgentConfig(agent="custom", install_skill=False)

    with pytest.raises(agent_config.AgentConfigError, match="No launch command"):
        agent_config.build_launch_argv(config, resolve=False)


def test_custom_agent_resolves_absolute_executable(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(tmp_path / "agent.json"))
    config = agent_config.AgentConfig(
        agent="custom",
        command=[sys.executable, "-V"],
        args=["--extra"],
        install_skill=False,
        preload_skill=False,
    )

    argv = agent_config.build_launch_argv(config, resolve=True)

    assert argv[0] == str(Path(sys.executable).resolve())
    assert argv[1:] == ["-V", "--extra"]


def test_custom_agent_wraps_windows_scripts(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(tmp_path / "agent.json"))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)
    cmd = tmp_path / "agent.cmd"
    cmd.write_text("@echo off\n", encoding="utf-8")
    ps1 = tmp_path / "agent.ps1"
    ps1.write_text("exit 0\n", encoding="utf-8")

    cmd_config = agent_config.AgentConfig(
        agent="custom",
        command=[str(cmd)],
        install_skill=False,
        preload_skill=False,
    )
    ps1_config = agent_config.AgentConfig(
        agent="custom",
        command=[str(ps1)],
        install_skill=False,
        preload_skill=False,
    )

    assert agent_config.build_launch_argv(cmd_config) == [
        os.environ.get("COMSPEC", "cmd.exe"),
        "/d",
        "/c",
        str(cmd.resolve()),
    ]
    assert agent_config.build_launch_argv(ps1_config) == [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(ps1.resolve()),
    ]


def test_cli_set_and_show(monkeypatch, tmp_path, capsys) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(tmp_path / "agent.json"))
    monkeypatch.setenv("HERMES_HOME", str(tmp_path / "hermes"))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)

    assert agent_config.main(["set", "hermes"]) == 0
    assert agent_config.load_agent_config().agent == "hermes"
    capsys.readouterr()

    assert agent_config.main(["show", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["ok"] is True
    assert payload["agent"] == "hermes"
    assert payload["command"] == ["hermes", "-s", "office"]


def test_cli_resolve_accepts_json(monkeypatch, tmp_path, capsys) -> None:
    path = tmp_path / "agent.json"
    agent_config.save_agent_config(
        agent_config.AgentConfig(
            agent="custom",
            command=[sys.executable],
            install_skill=False,
        ),
        path,
    )
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(path))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)

    assert agent_config.main(["resolve", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["ok"] is True
    assert payload["launch_argv"] == [str(Path(sys.executable).resolve())]


def test_cli_install_skill_overrides_disabled_custom(
    monkeypatch, tmp_path, capsys
) -> None:
    source = write_skill(tmp_path)
    path = tmp_path / "agent.json"
    agent_config.save_agent_config(
        agent_config.AgentConfig(
            agent="custom",
            command=[sys.executable],
            install_skill=False,
        ),
        path,
    )
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(path))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)

    assert (
        agent_config.main(
            [
                "install-skill",
                "--agent",
                "codex",
                "--project-root",
                str(tmp_path),
                "--source",
                str(source),
            ]
        )
        == 0
    )

    target = tmp_path / ".agents" / "skills" / "office" / "SKILL.md"
    assert target.is_file()
    assert str(target) in capsys.readouterr().out


def test_cli_show_rejects_unknown_agent(monkeypatch, tmp_path, capsys) -> None:
    path = tmp_path / "agent.json"
    path.write_text(json.dumps({"agent": "missing"}), encoding="utf-8")
    monkeypatch.setenv("OFFICE_BRIDGE_AGENT_CONFIG", str(path))
    monkeypatch.delenv("OFFICE_BRIDGE_AGENT", raising=False)

    assert agent_config.main(["show"]) == 2
    assert "Unknown agent" in capsys.readouterr().err
