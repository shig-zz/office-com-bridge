from __future__ import annotations

import pytest

from office_bridge.hotkey import MOD_ALT, MOD_CONTROL, MOD_SHIFT, HotkeyError, parse_hotkey


def test_parse_ctrl_alt_o() -> None:
    modifiers, key = parse_hotkey("ctrl+alt+o")
    assert modifiers == MOD_CONTROL | MOD_ALT
    assert key == ord("O")


def test_parse_shift_f5() -> None:
    modifiers, key = parse_hotkey("shift+f5")
    assert modifiers == MOD_SHIFT
    assert key == 0x74


def test_function_key_without_modifier_is_allowed() -> None:
    modifiers, key = parse_hotkey("f9")
    assert modifiers == 0
    assert key == 0x78


def test_plain_letter_requires_modifier() -> None:
    with pytest.raises(HotkeyError):
        parse_hotkey("o")


def test_unknown_key_is_rejected() -> None:
    with pytest.raises(HotkeyError):
        parse_hotkey("ctrl+not-a-key")


def test_empty_hotkey_is_rejected() -> None:
    with pytest.raises(HotkeyError):
        parse_hotkey("   ")
