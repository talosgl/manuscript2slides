"""Tests for application entry point routing when calling `python manuscript2slides`
(which enters execution at __main__.py)."""

# tests/test_main.py

import pytest
import sys
from unittest.mock import patch, MagicMock


def test_main_routes_to_gui_by_default(monkeypatch: pytest.MonkeyPatch) -> None:
    """When no args are provided at call time, main should call run_gui()."""

    # Monkeypatch - set up the test environment
    monkeypatch.setattr(sys, "argv", ["manuscript2slides"])

    # Mock: Replace the functions we want to track/fake (so the test won't call the real versions)
    with (
        patch(
            "manuscript2slides.__main__.startup.initialize_application"
        ) as mock_startup,
        patch("manuscript2slides.gui.run") as mock_gui,
        patch("manuscript2slides.__main__.run_cli") as mock_cli,
    ):
        # Inside this block, run_gui is replaced with a MagicMock
        # You can check if it was called, how many times, with what args, etc.

        # Configure the mock to return something (real startup returns a logger)
        mock_startup.return_value = MagicMock()

        # Run the actual code (with our mock replacements/overrides)
        from manuscript2slides.__main__ import main

        main()

        mock_gui.assert_called_once()
        mock_cli.assert_not_called()

    # Outside the "with" block, run_gui is back to normal; calling it will call the *real* version defined in __main__.py


def test_main_routes_to_cli_with_flag(monkeypatch: pytest.MonkeyPatch) -> None:
    """When --cli flag is present, __main__.py should route program execution flow to run_cli()"""
    monkeypatch.setattr(sys, "argv", ["manuscript2slides", "--cli"])

    with (
        patch(
            "manuscript2slides.__main__.startup.initialize_application"
        ) as mock_startup,
        patch("manuscript2slides.gui.run") as mock_gui,
        patch("manuscript2slides.__main__.run_cli") as mock_cli,
    ):
        mock_startup.return_value = MagicMock()

        from manuscript2slides.__main__ import main

        main()

        mock_cli.assert_called_once()
        mock_gui.assert_not_called()


def test_main_logs_and_reraises_exceptions(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test that main logs exceptions and re-raises them"""

    monkeypatch.setattr(sys, "argv", ["manuscript2slides"])

    with (
        patch(
            "manuscript2slides.__main__.startup.initialize_application"
        ) as mock_startup,
        patch("manuscript2slides.gui.run") as mock_gui,
    ):
        mock_logger = MagicMock()
        mock_startup.return_value = mock_logger

        # Make run_gui raise an exception
        test_exception = RuntimeError("Something went wrong!")
        mock_gui.side_effect = test_exception

        from manuscript2slides.__main__ import main

        # Verify the exception is re-raised
        with pytest.raises(RuntimeError, match="Something went wrong!"):
            main()

        # Verify the exception was logged
        mock_logger.exception.assert_called_once_with(
            "Unhandled exception - program crashed."
        )
