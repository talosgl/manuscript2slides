# tests/test_utils.py
"""Tests for utility functions."""

import pytest

from manuscript2slides.internals import constants
from manuscript2slides.utils import get_debug_mode, str_to_bool


# region basic pytest confidence check
def test_basic_assertion() -> None:
    """Verify pytest works."""
    assert 1 + 1 == 2


def test_string_comparison() -> None:
    """Verify string assertions work."""
    result = "hello"
    assert result == "hello"


# endregion


# region str_to_bool tests
@pytest.mark.parametrize(
    "input_str,expected",
    [
        # True values
        ("true", True),
        ("TRUE", True),
        ("tRuE", True),
        ("t", True),
        ("1", True),
        ("y", True),
        ("yEs", True),
        ("  true", True),
        # False values
        ("false", False),
        ("FALSE", False),
        ("fALsE", False),
        ("no", False),
        ("n", False),
        ("n  ", False),
    ],
)
def test_str_to_bool_returns_expected(input_str: str, expected: bool) -> None:
    """Test all valid true/false string variations"""
    assert str_to_bool(input_str) == expected


@pytest.mark.parametrize(
    "invalid_str", ["invalid", "maybe", "-1", "2", "", " ", "yes   no"]
)
def test_str_to_bool_raises_error_for_invalid_strings(invalid_str: str) -> None:
    """Test that invalid strings raise ValueError."""
    with pytest.raises(ValueError):
        str_to_bool(invalid_str)


def test_str_to_bool_error_message_includes_invalid_value() -> None:
    """Test that error message shows what value was invalid."""
    with pytest.raises(ValueError, match="bob"):
        str_to_bool("bob")


def test_str_to_bool_handles_whitespace() -> None:
    """Test that we handle whitespace inside values by stripping them."""
    assert str_to_bool("true ") == True
    assert str_to_bool("\ttrue\t") == True
    assert str_to_bool("\nfalse\n") == False
    assert str_to_bool("  yes  ") == True


# endregion

# region get_debug_mode tests


def test_get_debug_mode_returns_true_when_env_var_is_true(
    clean_debug_env: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode returns True when env var is set to 'true'."""
    clean_debug_env.setenv("MANUSCRIPT2SLIDES_DEBUG", "true")  # Only affects this test
    result = get_debug_mode()
    assert result == True


def test_get_debug_mode_returns_false_when_env_var_is_false(
    clean_debug_env: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode returns False when env var is set to 'false'."""
    clean_debug_env.setenv("MANUSCRIPT2SLIDES_DEBUG", "false")
    result = get_debug_mode()
    assert result == False


# By putting clean_debug_env in the parameter list, pytest:
#   1. Runs the fixture (which deletes the env var)
#   2. Passes the return value to your test (the monkeypatch object)
#   3. You don't use it, but the side effect already happened (env var deleted)
def test_get_debug_mode_returns_default_when_env_var_not_set(
    clean_debug_env: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode returns default (and does not crash to raised exception) when env var is not set."""
    # clean_debug_env was alreadycalled when this func was called, and deleted the var
    assert get_debug_mode() == constants.DEBUG_MODE_DEFAULT


def test_get_debug_mode_returns_default_when_env_var_invalid(
    clean_debug_env: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode handles invalid env var gracefully."""
    clean_debug_env.setenv("MANUSCRIPT2SLIDES_DEBUG", "invalid_value")
    # should fall back to default and not crash
    assert get_debug_mode() == constants.DEBUG_MODE_DEFAULT


def test_get_debug_mode_logs_warning_for_invalid_env_var(
    clean_debug_env: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that get_debug_mode logs a warning when env var is invalid."""
    import logging

    clean_debug_env.setenv("MANUSCRIPT2SLIDES_DEBUG", "banana")

    with caplog.at_level(logging.WARNING):
        get_debug_mode()

    # Check that a warning was indeed logged
    assert "Invalid value for MANUSCRIPT2SLIDES_DEBUG" in caplog.text
    assert "banana" in caplog.text


# endregion
