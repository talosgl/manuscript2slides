"""Baseline Tests for the GUI"""

import logging
from collections.abc import Generator
from pathlib import Path
from typing import Any

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMessageBox
from pytestqt.qtbot import QtBot

from manuscript2slides.gui import MainWindow, QTextEditHandler

# region Test Fixtures


@pytest.fixture(autouse=True)
def cleanup_gui_logger() -> Generator[None, Any, None]:
    """Remove GUI log handlers after each test to prevent cross-test pollution."""
    yield
    # After test completes, remove any QTextEditHandler instances
    # that reference Qt widgets to prevent cross-test contamination
    logger = logging.getLogger("manuscript2slides")

    # We're not literally removing handlers, instead, we're re-assigning logger.handlers
    # to a filtered version of its old list, keeping only handlers that are
    # NOT QTextEditHandler instances
    logger.handlers = [
        h for h in logger.handlers if not isinstance(h, QTextEditHandler)
    ]


@pytest.fixture
def mock_logger(monkeypatch: pytest.MonkeyPatch) -> logging.Logger:
    """Mock the GUI logger to prevent problematic GUI LogViewer widget access."""
    mock_logger = logging.getLogger("mock")
    monkeypatch.setattr("manuscript2slides.gui.log", mock_logger)
    return mock_logger


@pytest.fixture
def mock_critical_dialog(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, str]]:
    """Mock a QMessageBox.critical to allow tests to verify it is being called as expected.

    Returns a list of (title, text) tuples representing dialog calls.
    """

    # Track QMessageBox.critical calls
    dialog_calls: list[tuple[str, str]] = []

    def mock_critical(
        _parent: object, title: str, text: str, *args: object, **kwargs: object
    ) -> QMessageBox.StandardButton:
        dialog_calls.append((title, text))
        return QMessageBox.StandardButton.Ok

    monkeypatch.setattr(QMessageBox, "critical", mock_critical)

    return dialog_calls


@pytest.fixture
def mock_question_dialog(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, str]]:
    """Mock BaseConversionTabPresenter._show_question_dialog to allow tests to verify it is being called as expected.

    Returns a list of (title, text) tuples representing dialog calls.
    """
    dialog_calls: list[tuple[str, str]] = []

    def mock_show_question_dialog(*_args: object, **_kwargs: object) -> bool:
        title = _kwargs.get("title", _args[1] if len(_args) > 1 else "")
        text = _kwargs.get("text", _args[2] if len(_args) > 2 else "")
        dialog_calls.append((str(title), str(text)))
        return False

    monkeypatch.setattr(
        "manuscript2slides.gui.BaseConversionTabPresenter._show_question_dialog",
        mock_show_question_dialog,
    )
    return dialog_calls


# endregion


# region GUI Initialization


class TestGUIStartup:
    """Smoke tests - verify basic GUI launches without crashing."""

    def test_main_window_launches(self, qtbot: QtBot) -> None:
        """Test that MainWindow can be created without crashing."""
        window = MainWindow()
        qtbot.addWidget(window)

        assert window is not None
        assert window.windowTitle() == "manuscript2slides"

    def test_all_tabs_exist(self, qtbot: QtBot) -> None:
        """Test that all expected tabs are present."""
        window = MainWindow()
        qtbot.addWidget(window)

        assert window.tabs.count() == 3
        assert window.tabs.tabText(0) == "DOCX → PPTX"
        assert window.tabs.tabText(1) == "PPTX → DOCX"
        assert window.tabs.tabText(2) == "DEMO"

    def test_presenters_created(self, qtbot: QtBot) -> None:
        """Test that all presenters are initialized."""
        window = MainWindow()
        qtbot.addWidget(window)

        assert window.d2p_tab_presenter is not None
        assert window.p2d_tab_presenter is not None
        assert window.demo_presenter is not None

    def test_tab_switching(self, qtbot: QtBot) -> None:
        """Test that we can switch between all tabs without errors."""
        window = MainWindow()
        qtbot.addWidget(window)

        for i in range(3):
            window.tabs.setCurrentIndex(i)
            assert window.tabs.currentIndex() == i


class TestWidgetExistence:
    """Test that all expected widgets exist."""

    def test_demo_tab_has_buttons(self, qtbot: QtBot) -> None:
        """Test that demo tab has all expected buttons."""
        window = MainWindow()
        qtbot.addWidget(window)
        demo_view = window.demo_tab_view

        assert demo_view.docx2pptx_btn is not None
        assert demo_view.pptx2docx_btn is not None
        assert demo_view.round_trip_btn is not None
        assert demo_view.load_demo_btn is not None

    def test_pptx2docx_tab_has_io_widgets(self, qtbot: QtBot) -> None:
        """Test that PPTX→DOCX tab has I/O widgets."""
        window = MainWindow()
        qtbot.addWidget(window)
        p2d_view = window.p2d_tab_view

        assert p2d_view.input_selector is not None
        assert p2d_view.convert_btn is not None
        assert p2d_view.save_btn is not None
        assert p2d_view.load_btn is not None
        assert p2d_view.convert_btn.text() == "Convert!"

    def test_docx2pptx_tab_has_io_widgets(self, qtbot: QtBot) -> None:
        """Test that DOCX→PPTX tab has I/O widgets."""
        window = MainWindow()
        qtbot.addWidget(window)
        d2p_view = window.d2p_tab_view

        assert d2p_view.input_selector is not None
        assert d2p_view.convert_btn is not None
        assert d2p_view.save_btn is not None
        assert d2p_view.load_btn is not None
        assert d2p_view.convert_btn.text() == "Convert!"


# endregion


# region User Interactions


class TestButtonInteractions:
    """Test basic button click interactions."""

    def test_demo_button_click_doesnt_crash(
        self, qtbot: QtBot, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Test that clicking load demo button doesn't crash."""
        window = MainWindow()
        qtbot.addWidget(window)

        # Mock the file dialog to avoid blocking
        def mock_get_open_filename(
            parent: object, caption: str, directory: str, filter: str
        ) -> tuple[str, str]:
            # Return empty string (user cancelled)
            return ("", "")

        monkeypatch.setattr(
            "manuscript2slides.gui.QFileDialog.getOpenFileName",
            mock_get_open_filename,
        )

        # Click button - shouldn't crash
        qtbot.mouseClick(window.demo_tab_view.load_demo_btn, Qt.MouseButton.LeftButton)
        # If we got here, no crash occurred

    def test_convert_button_disabled_initially(self, qtbot: QtBot) -> None:
        """Test that convert button starts disabled with no input."""
        window = MainWindow()
        qtbot.addWidget(window)

        # Convert button should be disabled on startup (no input selected)
        assert not window.d2p_tab_view.convert_btn.isEnabled()
        assert not window.p2d_tab_view.convert_btn.isEnabled()

    def test_convert_button_enables_with_valid_file(
        self, qtbot: QtBot, tmp_path: Path
    ) -> None:
        """Test that convert button enables when valid file is selected."""
        window = MainWindow()
        qtbot.addWidget(window)

        # Create a dummy .docx file
        test_file = tmp_path / "test.docx"
        test_file.write_text("dummy content")

        # Set the path in the DOCX→PPTX tab
        window.d2p_tab_view.input_selector.set_path(str(test_file))

        # Wait for the button to become enabled, or 1000ms (whichever happens first)
        qtbot.waitUntil(
            lambda: window.d2p_tab_view.convert_btn.isEnabled(), timeout=1000
        )

        # Button should now be enabled
        assert window.d2p_tab_view.convert_btn.isEnabled()


# endregion


# region Validation Logic


class TestValidationLogic:
    """Test validation logic without full GUI lifecycle issues."""

    def test_docx2pptx_rejects_nonexistent_file(
        self,
        qtbot: QtBot,
        mock_logger: logging.Logger,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that validation rejects nonexistent input files."""
        # Create view and presenter without full MainWindow
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView
        from manuscript2slides.internals.define_config import UserConfig

        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Track QMessageBox.critical calls
        dialog_calls = mock_critical_dialog

        # Test with nonexistent file
        cfg = UserConfig()
        cfg.input_docx = Path("nonexistent.docx")

        result = presenter._validate_input(cfg)

        # Validation should fail
        assert result is False
        # Should have shown error dialog
        assert len(dialog_calls) == 1
        _title, text = dialog_calls[0]
        assert "not found" in text.lower() or "does not exist" in text.lower()

    def test_docx2pptx_rejects_wrong_file_type(
        self,
        qtbot: QtBot,
        tmp_path: Path,
        mock_logger: logging.Logger,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that validation rejects wrong file types."""
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView
        from manuscript2slides.internals.define_config import UserConfig

        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Track QMessageBox.critical calls
        dialog_calls = mock_critical_dialog

        # Create a .txt file instead of .docx
        wrong_file = tmp_path / "test.txt"
        wrong_file.write_text("not a docx")

        cfg = UserConfig()
        cfg.input_docx = wrong_file

        result = presenter._validate_input(cfg)

        # Validation should fail
        assert result is False
        # Should have shown error dialog about file type
        assert len(dialog_calls) == 1
        _title, text = dialog_calls[0]
        assert "invalid" in text.lower() or ".docx" in text.lower()

    def test_pptx2docx_rejects_nonexistent_file(
        self,
        qtbot: QtBot,
        mock_logger: logging.Logger,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that PPTX→DOCX validation rejects nonexistent files."""
        from manuscript2slides.gui import Pptx2DocxTabPresenter, Pptx2DocxTabView
        from manuscript2slides.internals.define_config import UserConfig

        view = Pptx2DocxTabView()
        qtbot.addWidget(view)
        presenter = Pptx2DocxTabPresenter(view)

        # Track QMessageBox.critical calls
        dialog_calls = mock_critical_dialog

        # Test with nonexistent file
        cfg = UserConfig()
        cfg.input_pptx = Path("nonexistent.pptx")

        result = presenter._validate_input(cfg)

        # Validation should fail
        assert result is False
        # Should have shown error dialog
        assert len(dialog_calls) == 1
        _title, text = dialog_calls[0]
        assert "not found" in text.lower() or "does not exist" in text.lower()

    def test_docx2pptx_rejects_dir_as_input(
        self,
        qtbot: QtBot,
        tmp_path: Path,
        mock_logger: logging.Logger,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that validation fails if a directory, instead of a file, is passed as input."""
        # Create View and Presenter
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView
        from manuscript2slides.internals.define_config import UserConfig

        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Track QMessageBox.critical calls
        dialog_calls = mock_critical_dialog

        # Create a directory (not a file)
        test_dir = tmp_path / "not_a_file"
        test_dir.mkdir()

        # Test validation
        cfg = UserConfig()
        cfg.input_docx = test_dir

        result = presenter._validate_input(cfg)

        assert result is False
        # Should have shown error dialog about file type
        assert len(dialog_calls) == 1
        title, text = dialog_calls[0]
        assert "invalid" in text.lower() or ".docx" in text.lower()


# endregion


# region Error Handling


class TestErrorHandling:
    """Test that exceptions during conversion are caught and presented to users."""

    def test_conversion_error_shows_dialog_and_logs(
        self,
        qtbot: QtBot,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        mock_logger: logging.Logger,
        mock_question_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that pipeline exceptions show error dialog and don't crash."""
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView

        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Track dialog calls
        dialog_calls = mock_question_dialog

        # Mock run_pipeline to raise an exception
        def mock_run_pipeline(_cfg: object) -> None:
            raise ValueError("Simulated conversion error")

        monkeypatch.setattr("manuscript2slides.gui.run_pipeline", mock_run_pipeline)

        # Create a valid input file (to pass validation)
        test_file = tmp_path / "test.docx"
        test_file.write_text("dummy content")
        view.input_selector.set_path(str(test_file))
        qtbot.waitUntil(lambda: view.convert_btn.isEnabled(), timeout=1000)

        # Trigger conversion (which will fail)
        presenter.on_convert_click()

        # Wait for worker thread to finish and process error
        # The thread emits an error signal which triggers _on_conversion_error
        qtbot.waitUntil(
            lambda: len(dialog_calls) >= 1
            and view.convert_btn.isEnabled(),  # wait until both expressions evaluate as true
            timeout=2000,
        )

        # Verify error dialog was shown
        assert len(dialog_calls) >= 1
        title, _text = dialog_calls[0]
        assert "conversion failed" in title.lower() or "failed" in title.lower()

        # Verify buttons were re-enabled after error
        assert view.convert_btn.isEnabled()

    def test_force_error_button_shows_dialog(
        self, qtbot: QtBot, mock_question_dialog: list[tuple[str, str]]
    ) -> None:
        """Test that the demo tab's force error button shows error dialog."""
        window = MainWindow()
        qtbot.addWidget(window)

        # Track dialog calls
        dialog_calls = mock_question_dialog

        # Click the force error button
        qtbot.mouseClick(
            window.demo_tab_view.force_error_btn, Qt.MouseButton.LeftButton
        )

        # Wait for worker thread to process error (Wait for dialog to be shown AND button to re-enable)
        qtbot.waitUntil(
            lambda: len(dialog_calls) >= 1
            and window.demo_tab_view.force_error_btn.isEnabled(),
            timeout=2000,
        )

        # Verify error dialog was shown
        assert len(dialog_calls) >= 1
        title, _text = dialog_calls[0]
        assert "conversion failed" in title.lower() or "failed" in title.lower()

        # Verify app is still responsive (buttons enabled)
        assert window.demo_tab_view.force_error_btn.isEnabled()

    def test_invalid_config_load_shows_error(
        self,
        qtbot: QtBot,
        tmp_path: Path,
        mock_logger: logging.Logger,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that loading malformed config shows error dialog."""
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView

        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Track QMessageBox.critical calls
        dialog_calls = mock_critical_dialog

        # Create a malformed TOML file
        bad_config = tmp_path / "bad_config.toml"
        bad_config.write_text("this is not valid TOML [[[[")

        # Try to load the bad config
        result = presenter._load_config(bad_config)

        # Should return None (failed to load)
        assert result is None

        # Should have shown error dialog
        assert len(dialog_calls) == 1
        title, text = dialog_calls[0]
        # Verify it mentions config loading failure
        assert "config" in text.lower() or "load" in text.lower()

    def test_worker_thread_exception_doesnt_crash(
        self,
        qtbot: QtBot,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        mock_question_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that exceptions in worker thread are caught and handled."""
        from manuscript2slides.gui import Pptx2DocxTabPresenter, Pptx2DocxTabView

        view = Pptx2DocxTabView()
        qtbot.addWidget(view)
        presenter = Pptx2DocxTabPresenter(view)

        # Track dialog calls
        dialog_calls = mock_question_dialog

        # Mock run_pipeline to raise RuntimeError (simulating deeper error)
        def mock_run_pipeline(_cfg: object) -> None:
            raise RuntimeError("Worker thread failure")

        monkeypatch.setattr("manuscript2slides.gui.run_pipeline", mock_run_pipeline)

        # Create valid input
        test_file = tmp_path / "test.pptx"
        test_file.write_text("dummy")
        view.input_selector.set_path(str(test_file))
        qtbot.waitUntil(lambda: view.convert_btn.isEnabled(), timeout=1000)

        # Trigger conversion
        presenter.on_convert_click()

        # Wait for worker to fail
        qtbot.waitUntil(
            lambda: len(dialog_calls) >= 1 and view.convert_btn.isEnabled(),
            timeout=2000,
        )

        # Verify error was handled (dialog shown)
        assert len(dialog_calls) >= 1

        # Verify buttons re-enabled (app recovered)
        assert view.convert_btn.isEnabled()

        # Verify worker thread cleaned up properly
        # (if it didn't, subsequent conversions would fail)
        assert presenter.worker_thread is not None  # Should still exist but be stopped


# endregion


# region Logging


class TestLogging:
    """Tests for logging behavior - separate from validation/error handling tests.

    These tests verify that log messages are written correctly.
    They use caplog instead of mocking the logger.
    """

    def test_validation_logs_error_for_nonexistent_file(
        self,
        qtbot: QtBot,
        caplog: pytest.LogCaptureFixture,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that nonexistent file validation logs the correct error message."""
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView
        from manuscript2slides.internals.define_config import UserConfig

        # Create view and presenter
        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Mock the dialog (still block GUI)
        dialog_calls = mock_critical_dialog

        # DON'T mock the logger - let caplog capture it!
        # (This is the key difference from validation tests)

        # Test validation with logging enabled
        cfg = UserConfig()
        cfg.input_docx = Path("nonexistent.docx")

        with caplog.at_level(logging.ERROR):
            result = presenter._validate_input(cfg)

        # Verify validation failed
        assert result is False

        # Verify logging happened with helpful error message
        assert any(
            r.levelname == "ERROR"
            and "does not exist" in r.message
            and str(cfg.input_docx) in r.message
            for r in caplog.records
        )

    def test_validation_logs_error_for_wrong_file_type(
        self,
        qtbot: QtBot,
        caplog: pytest.LogCaptureFixture,
        tmp_path: Path,
        mock_critical_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that wrong file type validation logs the correct error message."""
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView
        from manuscript2slides.internals.define_config import UserConfig

        # Create view and presenter
        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Mock the dialog
        dialog_calls = mock_critical_dialog

        # Create a file with wrong extension
        wrong_file = tmp_path / "test.txt"
        wrong_file.write_text("dummy content")

        # Test validation with logging enabled
        cfg = UserConfig()
        cfg.input_docx = wrong_file

        with caplog.at_level(logging.ERROR):
            result = presenter._validate_input(cfg)

        # Verify validation failed
        assert result is False

        # Verify logging happened correctly
        assert any(
            r.levelname == "ERROR"
            and (
                "invalid" in r.message.lower()
                or "must be a .docx file" in r.message.lower()
            )
            for r in caplog.records
        )

    def test_conversion_error_logs_exception(
        self,
        qtbot: QtBot,
        monkeypatch: pytest.MonkeyPatch,
        caplog: pytest.LogCaptureFixture,
        tmp_path: Path,
        mock_question_dialog: list[tuple[str, str]],
    ) -> None:
        """Test that conversion errors are logged with exception details."""
        from manuscript2slides.gui import Docx2PptxTabPresenter, Docx2PptxTabView

        # Create view and presenter
        view = Docx2PptxTabView()
        qtbot.addWidget(view)
        presenter = Docx2PptxTabPresenter(view)

        # Mock the presenter's dialog method
        dialog_calls = mock_question_dialog

        # Mock run_pipeline to raise exception
        def mock_run_pipeline(_cfg: object) -> None:
            raise ValueError("Simulated conversion error")

        monkeypatch.setattr("manuscript2slides.gui.run_pipeline", mock_run_pipeline)

        # Create valid input (passes validation!)
        test_file = tmp_path / "test.docx"
        test_file.write_text("dummy content")
        view.input_selector.set_path(str(test_file))
        qtbot.waitUntil(lambda: view.convert_btn.isEnabled(), timeout=2000)

        # Trigger conversion with logging capture
        with caplog.at_level(logging.ERROR):
            presenter.on_convert_click()

            # Wait for worker thread to fail
            qtbot.waitUntil(
                lambda: len(dialog_calls) >= 1 and view.convert_btn.isEnabled(),
                timeout=2000,
            )

        # Verify error was logged
        assert len(caplog.records) >= 1
        error_log = next((r for r in caplog.records if r.levelname == "ERROR"), None)
        assert error_log is not None
        assert (
            "conversion failed" in error_log.message.lower()
            or "error" in error_log.message.lower()
        )


# endregion


# region UI to Config


class TestUIToConfig:
    """Test that ui_to_config handles edge cases correctly."""

    def test_no_selection_converts_to_none_in_pptx2docx_ui_to_config(
        self, qtbot: QtBot
    ) -> None:
        """Test that 'No Selection' in input field becomes None in config, not a Path."""
        from manuscript2slides.internals.define_config import UserConfig

        window = MainWindow()
        qtbot.addWidget(window)

        # Get the PPTX→DOCX presenter and view
        presenter = window.p2d_tab_presenter
        view = window.p2d_tab_view

        # Simulate "No Selection" being displayed (happens when loading config with None)
        view.input_selector.set_path("No Selection")

        # Create a config and sync the view to it
        cfg = UserConfig()
        cfg = presenter.ui_to_config(cfg)

        # Verify config has None, not Path("No Selection")
        assert cfg.input_pptx is None

    def test_no_selection_converts_to_none_in_docx2pptx_ui_to_config(
        self, qtbot: QtBot
    ) -> None:
        """Test that 'No Selection' in input field becomes None in config for DOCX→PPTX."""
        from manuscript2slides.internals.define_config import UserConfig

        window = MainWindow()
        qtbot.addWidget(window)

        # Get the DOCX→PPTX presenter and view
        presenter = window.d2p_tab_presenter
        view = window.d2p_tab_view

        # Simulate "No Selection" being displayed
        view.input_selector.set_path("No Selection")

        # Create a config and sync the view to it
        cfg = UserConfig()
        cfg = presenter.ui_to_config(cfg)

        # Verify config has None, not Path("No Selection")
        assert cfg.input_docx is None

    def test_no_selection_for_template_and_output_paths_in_pptx2docx(
        self, qtbot: QtBot
    ) -> None:
        """Test that 'No Selection' for template and output paths becomes None."""
        from manuscript2slides.internals.define_config import UserConfig

        window = MainWindow()
        qtbot.addWidget(window)

        presenter = window.p2d_tab_presenter
        view = window.p2d_tab_view

        # Set all optional paths to "No Selection"
        view.output_selector.set_path("No Selection")
        view.template_selector.set_path("No Selection")

        # Sync to config
        cfg = UserConfig()
        cfg = presenter.ui_to_config(cfg)

        # Verify all are None
        assert cfg.output_folder is None
        assert cfg.template_docx is None

    def test_no_selection_for_template_and_output_paths_in_docx2pptx(
        self, qtbot: QtBot
    ) -> None:
        """Test that 'No Selection' for template and output paths becomes None."""
        from manuscript2slides.internals.define_config import UserConfig

        window = MainWindow()
        qtbot.addWidget(window)

        presenter = window.d2p_tab_presenter
        view = window.d2p_tab_view

        # Set all optional paths to "No Selection"
        view.output_selector.set_path("No Selection")
        view.template_selector.set_path("No Selection")

        # Sync to config
        cfg = UserConfig()
        cfg = presenter.ui_to_config(cfg)

        # Verify all are None
        assert cfg.output_folder is None
        assert cfg.template_pptx is None

    def test_config_roundtrip_preserves_none_values_pptx2docx(
        self, qtbot: QtBot
    ) -> None:
        """Test that None values survive a config→UI→config round trip."""
        from manuscript2slides.internals.define_config import UserConfig

        window = MainWindow()
        qtbot.addWidget(window)

        presenter = window.p2d_tab_presenter
        view = window.p2d_tab_view

        # Create a config with None values
        original_cfg = UserConfig()
        original_cfg.input_pptx = None
        original_cfg.output_folder = None
        original_cfg.template_docx = None
        original_cfg.range_start = None
        original_cfg.range_end = None

        # Round trip: config → UI → config
        view.config_to_ui(original_cfg)
        new_cfg = UserConfig()
        new_cfg = presenter.ui_to_config(new_cfg)

        # Verify None values are preserved
        assert new_cfg.input_pptx is None
        assert new_cfg.output_folder is None
        assert new_cfg.template_docx is None
        assert new_cfg.range_start is None
        assert new_cfg.range_end is None

    def test_config_roundtrip_preserves_none_values_docx2pptx(
        self, qtbot: QtBot
    ) -> None:
        """Test that None values survive a config→UI→config round trip."""
        from manuscript2slides.internals.define_config import UserConfig

        window = MainWindow()
        qtbot.addWidget(window)

        presenter = window.d2p_tab_presenter
        view = window.d2p_tab_view

        # Create a config with None values
        original_cfg = UserConfig()
        original_cfg.input_docx = None
        original_cfg.output_folder = None
        original_cfg.template_pptx = None
        original_cfg.range_start = None
        original_cfg.range_end = None

        # Round trip: config → UI → config
        view.config_to_ui(original_cfg)
        new_cfg = UserConfig()
        new_cfg = presenter.ui_to_config(new_cfg)

        # Verify None values are preserved
        assert new_cfg.input_docx is None
        assert new_cfg.output_folder is None
        assert new_cfg.template_pptx is None
        assert new_cfg.range_start is None
        assert new_cfg.range_end is None


# endregion
