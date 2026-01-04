"""Tests for CLI argument parsing and config building."""

import sys
from pathlib import Path
from unittest.mock import Mock

import pytest

from manuscript2slides.cli import (
    _validate_args_match_config,
    build_config_from_args,
    parse_args,
)
from manuscript2slides.internals.define_config import ChunkType, UserConfig


# region TestParseArgs
class TestParseArgs:
    """Test that parse_args stores the values we expect."""

    @pytest.mark.parametrize(
        argnames="arg_dest_name,cli_flag,expected",
        argvalues=[
            ("experimental_formatting_on", "--experimental-formatting", True),
            ("experimental_formatting_on", "--no-experimental-formatting", False),
            ("display_comments", "--display-comments", True),
            ("display_comments", "--no-display-comments", False),
            ("comments_sort_by_date", "--comments-sort-by-date", True),
            ("comments_sort_by_date", "--no-comments-sort-by-date", False),
            ("comments_keep_author_and_date", "--comments-keep-author-and-date", True),
            (
                "comments_keep_author_and_date",
                "--no-comments-keep-author-and-date",
                False,
            ),
            ("display_footnotes", "--display-footnotes", True),
            ("display_footnotes", "--no-display-footnotes", False),
            ("display_endnotes", "--display-endnotes", True),
            ("display_endnotes", "--no-display-endnotes", False),
            ("preserve_docx_metadata_in_speaker_notes", "--preserve-metadata", True),
            (
                "preserve_docx_metadata_in_speaker_notes",
                "--no-preserve-metadata",
                False,
            ),
        ],
    )
    def test_cli_boolean_flags_set_correctly_when_provided(
        self,
        monkeypatch: pytest.MonkeyPatch,
        arg_dest_name: str,
        cli_flag: str,
        expected: bool,
    ) -> None:
        """Test that boolean flags set correct True/False values when they are provided explicitly."""
        monkeypatch.setattr(sys, "argv", ["manuscript2slides.cli", cli_flag])
        args = parse_args()
        assert getattr(args, arg_dest_name) == expected

    @pytest.mark.parametrize(
        argnames="arg_dest_name",
        argvalues=[
            "experimental_formatting_on",
            "display_comments",
            "comments_sort_by_date",
            "comments_keep_author_and_date",
            "display_footnotes",
            "display_endnotes",
            "preserve_docx_metadata_in_speaker_notes",
        ],
    )
    def test_boolean_flags_default_to_none_when_not_provided(
        self, monkeypatch: pytest.MonkeyPatch, arg_dest_name: str
    ) -> None:
        """
        Test that boolean flags are None when not provided. For example,
        When neither flag ("--experimental-formatting" nor "--no-experimental-formatting")
        is provided, the dest ("experimental_formatting_on") should be set to None.
        """
        monkeypatch.setattr(
            sys,
            "argv",
            [
                "manuscript2slides.cli",
                "--demo-docx2pptx",
            ],
        )
        args = parse_args()
        assert getattr(args, arg_dest_name) is None

    @pytest.mark.parametrize(
        argnames="arg_dest_name",
        argvalues=[
            "demo_docx2pptx",
            "demo_pptx2docx",
            "demo_round_trip",
        ],
    )
    def test_demo_flag_false_when_not_provided(
        self, monkeypatch: pytest.MonkeyPatch, arg_dest_name: str
    ) -> None:
        """Demo flags should evaluate to False when not provided."""
        monkeypatch.setattr(
            sys,
            "argv",
            [
                "manuscript2slides.cli",
                "--input-docx",
                "dummy.docx",
            ],
        )
        args = parse_args()
        assert getattr(args, arg_dest_name) is False

    @pytest.mark.parametrize(
        argnames="cli_flag,arg_dest_name",
        argvalues=[
            ("--demo-docx2pptx", "demo_docx2pptx"),
            ("--demo_pptx2docx", "demo_pptx2docx"),
            ("--demo_round_trip", "demo_round_trip"),
        ],
    )
    def test_demo_flag_is_true_when_provided(
        self, monkeypatch: pytest.MonkeyPatch, cli_flag: str, arg_dest_name: str
    ) -> None:
        """Demo flags should cause the arg destination name evaluate to True if they are provided."""
        monkeypatch.setattr(sys, "argv", ["manuscript2slides.cli", cli_flag])
        args = parse_args()
        assert getattr(args, arg_dest_name) is True


# endregion


# region TestBuildConfigFromArgs
class TestBuildConfigFromArgs:
    """Test that build_config_from_args respects CLI > config file > defaults priority."""

    def test_cli_uses_defaults_for_unspecified_fields(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Config defaults should be used when no overriding flags were provided."""
        monkeypatch.setattr(
            sys, "argv", ["manuscript2slides.cli", "--input-docx", "dummy.docx"]
        )  # Don't fail this particular test because no input was provided
        args = parse_args()
        cfg = build_config_from_args(args)

        default_cfg = UserConfig(
            input_docx=Path("dummy.docx")
        )  # This doesn't really need to match but it's confusing when reading the test if it doesn't.

        assert cfg.experimental_formatting_on == default_cfg.experimental_formatting_on
        assert cfg.chunk_type == default_cfg.chunk_type
        assert cfg.template_docx == default_cfg.template_docx

    @pytest.mark.parametrize(
        argnames="arg_dest_name,cli_flag,expected",
        argvalues=[
            ("experimental_formatting_on", "--experimental-formatting", True),
            ("experimental_formatting_on", "--no-experimental-formatting", False),
            ("display_comments", "--display-comments", True),
            ("display_comments", "--no-display-comments", False),
            ("comments_sort_by_date", "--comments-sort-by-date", True),
            ("comments_sort_by_date", "--no-comments-sort-by-date", False),
            ("display_footnotes", "--display-footnotes", True),
            ("display_footnotes", "--no-display-footnotes", False),
            ("display_endnotes", "--display-endnotes", True),
            ("display_endnotes", "--no-display-endnotes", False),
            ("preserve_docx_metadata_in_speaker_notes", "--preserve-metadata", True),
            (
                "preserve_docx_metadata_in_speaker_notes",
                "--no-preserve-metadata",
                False,
            ),
        ],
    )
    def test_build_config_from_args_respects_cli_boolean_overrides(
        self,
        monkeypatch: pytest.MonkeyPatch,
        arg_dest_name: str,
        cli_flag: str,
        expected: bool,
    ) -> None:
        """Setting an option to True or False in the CLI should override config default."""
        monkeypatch.setattr(
            sys,
            "argv",
            ["manuscript2slides.cli", "--input-docx", "dummy.docx", cli_flag],
        )
        args = parse_args()
        cfg = build_config_from_args(args)

        assert getattr(cfg, arg_dest_name) == expected

    @pytest.mark.parametrize(
        argnames="cli_value,expected_enum",
        argvalues=[
            ("paragraph", ChunkType.PARAGRAPH),
            ("page", ChunkType.PAGE),
            ("heading", ChunkType.HEADING_FLAT),
            ("heading_flat", ChunkType.HEADING_FLAT),
            ("heading_nested", ChunkType.HEADING_NESTED),
        ],
    )
    def test_build_config_from_args_chunk_type_enum_conversion(
        self, monkeypatch: pytest.MonkeyPatch, cli_value: str, expected_enum: ChunkType
    ) -> None:
        """Test that chunk_type string converts to enum."""
        monkeypatch.setattr(
            sys,
            "argv",
            [
                "manuscript2slides.cli",
                "--input-docx",
                "dummy.docx",
                "--chunk-type",
                cli_value,
            ],
        )
        args = parse_args()
        cfg = build_config_from_args(args)

        assert cfg.chunk_type == expected_enum


# endregion


# region test _validate_args_match_config
def test_validate_args_match_config_returns_early_when_debug_mode_false(
    clean_debug_env: pytest.MonkeyPatch,
) -> None:
    """Verify that users will never mistakenly get app crashes because of config field/arg name misalignment."""

    # Arrange: ensure get_debug_mode() will return FALSE when called inside _validate_args_match_config() (env var set to false)
    clean_debug_env.setenv("MANUSCRIPT2SLIDES_DEBUG", "false")

    # Arrange: prepare bad data to pass into the validator
    fake_parser = Mock()
    fake_action = Mock()
    fake_action.dest = "fake_bad_field"  # Anything not in UserConfig class
    fake_parser._actions = [fake_action]

    # Action: Pass in a mock with obviously bad dest fields that WOULD cause an exception if debug mode was set to true
    _validate_args_match_config(fake_parser)

    # Test will fail if the above raises an exception; no additional assert needed.


def test_validate_args_match_config_raises_exception_when_debug_mode_true_and_bad_data(
    clean_debug_env: pytest.MonkeyPatch,
) -> None:
    """Verify that developers DO get an error in debug mode when there is config field/arg name misalignment."""

    # Arrange: ensure get_debug_mode() will return TRUE when called inside _validate_args_match_config() (env var set to false)
    clean_debug_env.setenv("MANUSCRIPT2SLIDES_DEBUG", "true")

    # Arrange: prepare bad data to pass into the validator
    fake_parser = Mock()
    fake_action = Mock()
    fake_action.dest = "fake_bad_field"  # Anything not in UserConfig class
    fake_parser._actions = [fake_action]

    # Action: call _validate_args_match_config with bad data
    # Should raise RuntimeError in debug mode
    with pytest.raises(RuntimeError):
        _validate_args_match_config(fake_parser)


def test_real_parser_passes_validation_in_debug_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify that the real CLI parser configuration passes validation in debug mode.

    This test ensures that all CLI arguments are properly configured and either:
    1. Have corresponding UserConfig fields, or
    2. Are listed in the excluded_args set (like --help, --version, --config, etc.)

    If this test fails, it means you either:
    - Added a CLI arg without adding it to UserConfig
    - Added a CLI arg without excluding it in _validate_args_match_config
    """
    # Arrange: enable debug mode so validation actually runs
    monkeypatch.setenv("MANUSCRIPT2SLIDES_DEBUG", "true")

    # Set up argv with a valid demo flag so parse_args() doesn't error out
    monkeypatch.setattr(sys, "argv", ["manuscript2slides.cli", "--demo-docx2pptx"])

    # Action: parse_args() internally calls _validate_args_match_config()
    # This should NOT raise an exception if everything is configured correctly
    args = parse_args()

    # Assert: if we got here without an exception, validation passed
    assert args is not None


# endregion
