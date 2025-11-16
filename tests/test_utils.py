# tests/test_utils.py
"""Tests for utility functions."""
import pytest

from manuscript2slides.internals import constants
from manuscript2slides.utils import str_to_bool, get_debug_mode


def test_str_to_bool_returns_true_for_valid_true_strings() -> None:
    """Test that the various 'true' strings return True."""

    # Most common case
    assert (
        str_to_bool("true") == True
    )  # "I assert these match and if I'm wrong, crash the program."

    # test case sensitivity
    assert str_to_bool("TRUE") == True
    assert str_to_bool("tRuE") == True

    # Test other valid true values
    assert str_to_bool("t") == True
    assert str_to_bool("1") == True
    assert str_to_bool("yes") == True
    assert str_to_bool("y") == True
    # Test multiple cases in one test when they're closely related


def test_str_to_bool_returns_true_for_valid_false_strings() -> None:
    """Test that the various 'false' strings return True."""

    assert str_to_bool("false") == False
    assert str_to_bool("FALSE") == False
    assert str_to_bool("fALsE") == False

    assert str_to_bool("f") == False
    assert str_to_bool("0") == False
    assert str_to_bool("n") == False
    assert str_to_bool("No") == False


def test_str_to_bool_raises_error_for_invalid_strings() -> None:
    """Test that invalid strings raise ValueError."""

    with pytest.raises(ValueError):
        str_to_bool("invalid")

    with pytest.raises(ValueError):
        str_to_bool("maybe")

    with pytest.raises(ValueError):
        str_to_bool("-1")

    with pytest.raises(ValueError):
        str_to_bool("2")

    with pytest.raises(ValueError):
        str_to_bool("")

    with pytest.raises(ValueError):
        str_to_bool(" ")

    with pytest.raises(ValueError):
        str_to_bool("yes  no ")


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


def test_get_debug_mode_returns_true_when_env_var_is_true(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode returns True when env var is set to 'true'."""
    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "true")
    result = get_debug_mode()
    assert result == True


def test_get_debug_mode_returns_false_when_env_var_is_false(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode returns False when env var is set to 'false'."""
    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "false")
    result = get_debug_mode()
    assert result == False


# Not sure how useful this is to test since we already tested case sensitivity for the called function above?
def test_get_debug_mode_handles_case_insensitive_values(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode handles case variations."""
    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "TRUE")
    assert get_debug_mode() == True

    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "FALSE")
    assert get_debug_mode() == False


def test_get_debug_mode_returns_default_when_env_var_not_set(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode returns default (and does not crash to raised exception) when env var is not set."""
    monkeypatch.delenv("MANUSCRIPT2SLIDES_DEBUG", raising=False)
    # raising=False means "don't error if it doesn't exist"
    assert get_debug_mode() == constants.DEBUG_MODE_DEFAULT


def test_get_debug_mode_returns_default_when_env_var_invalid(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that get_debug_mode handles invalid env var gracefully."""
    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "invalid_value")
    # should fall back to default and not crash
    assert get_debug_mode() == constants.DEBUG_MODE_DEFAULT


def test_get_debug_mode_logs_warning_for_invalid_env_var(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that get_debug_mode logs a warning when env var is invalid."""
    import logging

    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "banana")

    with caplog.at_level(logging.WARNING):
        get_debug_mode()

    # Check that a warning was indeed logged
    assert "Invalid value for MANUSCRIPT2SLIDES_DEBUG" in caplog.text
    assert "banana" in caplog.text
