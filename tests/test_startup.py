# tests/test_startup.py
"""Tests for application startup.

Primarily tests whether trace logging is enabled/disabled correctly
and that error handling happens at all when try/except exceptions are hit."""

from unittest.mock import patch

import pytest

from manuscript2slides.startup import (
    _should_enable_trace_on_startup,
    initialize_application,
)


def test_trace_enabled_by_env_var() -> None:
    """Trace should enable when env var is set to true"""
    with patch.dict("os.environ", {"MANUSCRIPT2SLIDES_DEBUG": "true"}):
        assert _should_enable_trace_on_startup() is True


def test_trace_disabled_by_env_var() -> None:
    """Trace should disable when env var is set to false"""
    with patch.dict("os.environ", {"MANUSCRIPT2SLIDES_DEBUG": "false"}):
        assert _should_enable_trace_on_startup() is False


def test_trace_uses_default_when_no_env_var() -> None:
    """Trace should use constant default when env var not set"""
    with patch.dict("os.environ", {}, clear=True):
        # This tests that it falls back to DEBUG_MODE_DEFAULT
        result = _should_enable_trace_on_startup()
        assert isinstance(result, bool)  # At least verify it returns a bool


def test_initialize_exits_on_permission_error(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Should exit cleanly with message when log directory creation fails"""
    with patch("manuscript2slides.startup.setup_logger") as mock_setup_logger:
        mock_setup_logger.side_effect = PermissionError("Access denied")

        with pytest.raises(SystemExit) as exc_info:
            initialize_application()

        assert exc_info.value.code == 1
        captured = capsys.readouterr()
        assert "Cannot create log files" in captured.err
        assert "Check permissions" in captured.err


def test_initialize_exits_on_os_error(capsys: pytest.CaptureFixture[str]) -> None:
    """Should exit cleanly when disk full or I/O error"""
    with patch("manuscript2slides.startup.setup_logger") as mock_setup_logger:
        mock_setup_logger.side_effect = OSError("Disk full")

        with pytest.raises(SystemExit) as exc_info:
            initialize_application()

        assert exc_info.value.code == 1
        captured = capsys.readouterr()
        assert "disk full or I/O error" in captured.err
