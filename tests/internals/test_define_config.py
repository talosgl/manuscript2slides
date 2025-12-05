"""Tests for UserConfig class definition file and related items"""

import pytest
from manuscript2slides.internals.define_config import (
    ChunkType,
    UserConfig,
    PipelineDirection,
)
import sys
import logging
from pathlib import Path


# region test ChunkType enum
@pytest.mark.parametrize(
    argnames="input_string,expected_enum",
    argvalues=[
        # Standard values
        ("paragraph", ChunkType.PARAGRAPH),
        ("page", ChunkType.PAGE),
        ("heading_flat", ChunkType.HEADING_FLAT),
        ("heading_nested", ChunkType.HEADING_NESTED),
        # Alias
        ("heading", ChunkType.HEADING_FLAT),
        # Test normalization (case insensitive, strips whitespace)
        ("PARAGRAPH", ChunkType.PARAGRAPH),
        ("  page  ", ChunkType.PAGE),
        ("Heading_Flat", ChunkType.HEADING_FLAT),
    ],
)
def test_chunk_type_from_string(input_string: str, expected_enum: ChunkType) -> None:
    """Test ChunkType.from_string() with valid values, aliases, and normalization."""
    result = ChunkType.from_string(input_string)
    assert result == expected_enum


def test_chunk_type_from_string_invalid() -> None:
    """Test ChunkType.from_string() raises ValueError for invalid input."""
    with pytest.raises(ValueError, match="'invalid' is not a valid ChunkType"):
        ChunkType.from_string("invalid")


# endregion

# region UserConfig class


# region from_toml tests
# Case: Happy paths - # Provide several valid toml configurations and ensure there's no raises.
@pytest.mark.parametrize(
    argnames="mock_toml",
    argvalues=[
        # basic docx config
        """
    input_docx = "./sample_doc.docx"

# === Processing Options ===
chunk_type = "page" # Options: paragraph, page, heading_flat, heading_nested

# === Formatting ===
experimental_formatting_on = true

# === Annotations ===
preserve_docx_metadata_in_speaker_notes = true

    """,
        # basic pptx config
        """
    input_pptx = "./test_slides.pptx"
    template_docx = "./docx_template.docx"
""",
        # can add more cases below
        # ...
    ],
)
def test_from_toml_happy_paths(tmp_path: Path, mock_toml: str) -> None:
    """Test that known-good toml config files do not raise when parsed with from_toml."""
    toml_file = tmp_path / "config.toml"
    toml_file.write_text(mock_toml)

    # Test will fail if any errors are hit on this call
    test_cfg = UserConfig.from_toml(toml_file)

    # Case: if the file includes chunk_type as string, does it makes it out as an enum?
    if "chunk_type" in mock_toml:
        assert isinstance(test_cfg.chunk_type, ChunkType)


def test_from_toml_real_path(sample_config_toml: Path) -> None:
    """Ensure we load fine from a known good file."""
    UserConfig.from_toml(sample_config_toml)


@pytest.mark.parametrize(
    argnames="mock_toml",
    argvalues=[
        # Case: Mismatched/unclosed quotes
        """
        input_pptx = ./test_slides.pptx"
        template_docx = "./docx_template.docx"
        """,
        # Case: Missing quotes around string value
        """
        input_docx = ./sample_doc.docx
        """,
        # Case: Mixing quote types
        """
        input_docx = "./sample.docx'
        """,
        # Case: Unquoted path with special characters
        r"""
        input_docx = C:\Users\test.docx

        """,
    ],
)
def test_from_toml_bad_quotes(tmp_path: Path, mock_toml: str) -> None:
    """Ensure we raise with a helpful message when encountering toml syntax parsing errors from quotation marks."""
    toml_file = tmp_path / "config.toml"
    toml_file.write_text(mock_toml)

    with pytest.raises(ValueError, match="quote"):
        test_cfg = UserConfig.from_toml(toml_file)


def test_from_toml_raises_with_binary_file(path_to_empty_docx: Path) -> None:
    """Test that passing a non-TOML file (like .docx) raises appropriately."""
    with pytest.raises(ValueError, match="decode"):
        UserConfig.from_toml(path_to_empty_docx)


def test_from_toml_warns_if_empty(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Ensure we log a warning if ingested toml file is loaded as empty."""
    toml_file = tmp_path / "config.toml"
    toml_file.write_text("")

    with caplog.at_level(logging.WARNING):
        test_cfg = UserConfig.from_toml(toml_file)
        assert "Config toml file loaded as empty" in caplog.text
        assert "feedback" in caplog.text


def test_from_toml_raises_helpfully_if_path_not_exist(tmp_path: Path) -> None:
    """Ensure we raise if user passed in path to a file that doesn't exist"""

    with pytest.raises(FileNotFoundError, match="file not found"):
        test_cfg = UserConfig.from_toml(tmp_path / "fake.toml")


def test_from_toml_raises_helpfully_if_path_is_dir(tmp_path: Path) -> None:
    """Ensure we raise if user passed in path to a file that doesn't exist"""

    with pytest.raises(ValueError, match="is a directory"):
        test_cfg = UserConfig.from_toml(tmp_path)


# Cases: Ensure we raise if:
# Unable to parse/decode the [toml] file contents into a py dict
def test_from_toml_raises_if_cannot_parse_to_dict(tmp_path: Path) -> None:
    """Ensure we raise when encountering toml syntax parsing errors (generic)."""

    mock_toml = """
    What happens if this is just a bunch of normal text?
    """

    toml_file = tmp_path / "config.toml"
    toml_file.write_text(mock_toml)

    with pytest.raises(
        Exception
    ):  # Pretty weak check but this is kind of a bucket for anything we haven't thought of writing specific syntax tests for.
        test_cfg = UserConfig.from_toml(toml_file)


def test_from_toml_raises_helpfully_with_bad_chunk_type_data(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Ensure we raise with a helpful message if chunk_type has invalid value"""
    toml_file = tmp_path / "config.toml"
    toml_file.write_text(
        """
        input_docx = "./sample_doc.docx"
        chunk_type = "not_a_real_chunk_type"
        """
    )

    with pytest.raises(ValueError, match="Invalid chunk_type"):
        test_cfg = UserConfig.from_toml(toml_file)


def test_from_toml_unexpected_fields_warn_and_filter(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Test that unexpected fields in TOML raise ValueError."""
    toml_file = tmp_path / "config.toml"
    toml_file.write_text(
        """
        input_docx = "test.docx"
        chunk_type = "paragraph"
        invalid_field = "should not be here"
        another_bad_field = 123
        """
    )
    with caplog.at_level(logging.WARNING):
        test_cfg = UserConfig.from_toml(toml_file)
        assert "Ignoring unexpected fields" in caplog.text
        assert "Check for typos" in caplog.text
        assert "Filtering out unexpected fields" in caplog.text


def test_from_toml_does_normalize_paths(tmp_path: Path) -> None:
    """Test that paths with backslashes get normalized to forward slashes."""
    toml_file = tmp_path / "config.toml"
    toml_file.write_text('input_docx = "C:\\\\Users\\\\test\\\\file.docx"\n')

    cfg = UserConfig.from_toml(toml_file)

    assert cfg.input_docx == "C:/Users/test/file.docx"


# end region


# region direction property tests
@pytest.mark.parametrize(
    argnames="input_cfg,expected_result",
    argvalues=[
        # Valid input
        (UserConfig(input_docx="path/to/input.docx"), PipelineDirection.DOCX_TO_PPTX),
        (UserConfig(input_pptx="path/to/input.pptx"), PipelineDirection.PPTX_TO_DOCX),
    ],
)
def test_direction_property(
    input_cfg: UserConfig, expected_result: PipelineDirection
) -> None:
    """Test we get expected results for expected input"""
    result = input_cfg.direction
    print(expected_result, result)
    assert result == expected_result


def test_direction_property_raises_when_neither_input_path_set(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Ensure we raise if we can't set direction because neither input field was set.
    Not clear how we'd ever get to this situation in the current codebase, but, just in case.
    """
    test_cfg = UserConfig(input_docx=None, input_pptx=None)
    with pytest.raises(ValueError, match="no input"):
        result = test_cfg.direction
    assert "no input_docx or input_pptx path" in caplog.text


def test_direction_property_raises_when_both_inputs_set(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Ensure we raise if we we cannot set direction because both input fields were set."""
    test_cfg = UserConfig(input_docx="a_docx_path.docx", input_pptx="a_pptx_path.pptx")
    with pytest.raises(ValueError, match="too many inputs"):
        result = test_cfg.direction
    assert "Only 1 input can be specified" in caplog.text


# endregion

# region pri3 TODOs


# TODO: might be low value-to-effort, but, test all the get_*_path() ?
def test_get_input_docx_path_works(
    path_to_sample_docx_with_everything: Path,
) -> None:

    test_cfg = UserConfig(input_docx=str(path_to_sample_docx_with_everything))
    result = test_cfg.get_input_docx_file()

    assert (
        result is not None
    ), f"Test result is evaluating to None: probably it is not able to access the real test data from disk."

    # Compare resolved/absolute paths
    assert result.resolve() == path_to_sample_docx_with_everything.resolve()

    # Or just check the filename matches
    assert result.name == path_to_sample_docx_with_everything.name


def test_get_input_docx_file_returns_none_when_not_set() -> None:
    """Test that get_input_docx_file returns None when input_docx is not set."""
    test_cfg = UserConfig(input_pptx="something.pptx")  # Set other input instead
    result = test_cfg.get_input_docx_file()
    assert result is None


def test_get_template_pptx_path_returns_default_when_not_set() -> None:
    """Test that get_template_pptx_path falls back to default template."""
    test_cfg = UserConfig(input_docx="test.docx")  # Don't set template_pptx
    result = test_cfg.get_template_pptx_path()
    # Check it returns the default (not None)
    assert result is not None
    assert result.exists()  # Default should exist


def test_get_output_folder_returns_configured_path(tmp_path: Path) -> None:
    """Test that get_output_folder returns configured path when set."""
    test_cfg = UserConfig(input_docx="test.docx", output_folder=str(tmp_path))
    result = test_cfg.get_output_folder()
    assert result == tmp_path


def test_get_output_folder_returns_default_when_not_set() -> None:
    """Test that get_output_folder falls back to default output dir."""
    test_cfg = UserConfig(input_docx="test.docx")
    result = test_cfg.get_output_folder()
    assert result is not None  # Should return default, not None


@pytest.mark.parametrize(
    argnames="bool_name",
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
def test_enable_all_options(sample_d2p_cfg: UserConfig, bool_name: str) -> None:
    """Ensure enable_all_options really sets everything to true."""
    # Arrange: Set this specific bool to False
    setattr(sample_d2p_cfg, bool_name, False)

    # Act: Call the function
    sample_d2p_cfg.enable_all_options()

    # Assert: Check it's now True
    assert getattr(sample_d2p_cfg, bool_name) is True


# TODO: test save_toml
#    Case: Test which takes a config from memory, saves with .save_toml to tmp_path/file.toml, then calls load_toml.

"""
Suggestions:
Test the round-trip (create config → save → load → verify match)
Test that None values are filtered out of the saved TOML
Test that parent directories get created if they don't exist
Test that it raises when path is a directory
Test that enums are saved as strings in the TOML
"""


# TODO: test config_to_dict
"""
Suggestions:
Test that all fields get converted to the dict
Test that enums get converted to their string values (.value)
Test that paths get normalized
Test that None values are included in the dict (filtering happens in save_toml)
"""
# endregion


# region validate() tests
def test_validate_valid_config(sample_d2p_cfg: UserConfig) -> None:
    """Ensure validate() succeeds when passed a valid config"""
    cfg = sample_d2p_cfg
    cfg.validate()  # should not raise


@pytest.mark.parametrize(
    argnames="config_input,expected_exception,match_pattern",
    argvalues=[
        # Only 1 input file allowed
        (
            UserConfig(input_docx="file.docx", input_pptx="file.pptx"),
            ValueError,
            "Cannot specify both",
        ),
        # Must have at least one input set
        (
            UserConfig(input_docx=None, input_pptx=None),
            ValueError,
            "No input file",
        ),
        # Output folder cannot be empty string
        (
            UserConfig(input_docx="input.docx", output_folder=""),
            ValueError,
            "cannot be empty string; use None for default",
        ),
        # Range checks
        (
            UserConfig(input_docx="input.docx", range_start=5, range_end=2),
            ValueError,
            "range_start .* cannot be greater than range_end",
        ),
    ],
)
def test_validate_raises_when_passed_bad_data(
    config_input: UserConfig,
    expected_exception: type[Exception],
    match_pattern: str,
) -> None:
    """Ensure UserConfig.validate() catches bad data beyond type issues.
    Type/is-instance checks should be tested in the load_toml() testing.
    """
    with pytest.raises(expected_exception, match=match_pattern):
        config_input.validate()


def test_validate_catches_and_warns_for_input_and_template_passed_with_same_filetype(
    sample_d2p_cfg: UserConfig,
    path_to_empty_docx: Path,
    sample_p2d_cfg: UserConfig,
    path_to_empty_pptx: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Ensure we log a warning to the log if the user passes in both input_docx + template_docx or input_pptx + template_pptx, because we're going to ignore their template."""
    sample_d2p_cfg.template_docx = str(path_to_empty_docx)
    sample_p2d_cfg.template_pptx = str(path_to_empty_pptx)

    with caplog.at_level(logging.WARN):
        sample_d2p_cfg.validate()
        sample_p2d_cfg.validate()
    assert (
        "You provided a template_docx, but you passed an input_docx for the pipeline"
        in caplog.text
    )
    assert (
        "You provided a template_pptx, but you passed an input_pptx for the pipeline"
        in caplog.text
    )
    assert "will be ignored" in caplog.text


# endregion

# region test_validate_docx2pptx_pipeline_requirements


def test_validate_docx2pptx_pipeline_requirements_valid_config(
    sample_d2p_cfg: UserConfig,
) -> None:
    """Test the baseline happy path for doc2pptx pipeline validation."""
    sample_d2p_cfg.validate_docx2pptx_pipeline_requirements()  # should not raise


@pytest.mark.parametrize(
    argnames="config_input,expected_exception,match_pattern",
    argvalues=[
        # Cases:
        # Input path is None
        (UserConfig(input_docx=None), ValueError, "No input docx file specified"),
        # Input path does not exist on disk
        (
            UserConfig(input_docx="this_file_doesn't_exist.docx"),
            FileNotFoundError,
            "file not found",
        ),
    ],
)
def test_validate_docx2pptx_pipeline_requirements_raises_if_input_None_or_not_exist(
    config_input: UserConfig,
    expected_exception: type[Exception],
    match_pattern: str,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Ensure we raise with helpful error message/logging when input_docx is None or doesn't exist on disk."""

    with pytest.raises(expected_exception, match=match_pattern):
        config_input.validate_docx2pptx_pipeline_requirements()
    assert match_pattern in caplog.text


def test_validate_docx2pptx_pipeline_requirements_raises_if_input_docx_is_dir(
    tmp_path: Path,
) -> None:
    """Test we raise gracefully if a directory, instead of file, is passed to input_docx."""
    test_cfg = UserConfig(input_docx=str(tmp_path))
    with pytest.raises(expected_exception=ValueError, match="not a file"):
        test_cfg.validate_docx2pptx_pipeline_requirements()


def test_validate_docx2pptx_pipeline_requirements_raises_if_template_not_exist(
    sample_d2p_cfg: UserConfig,
) -> None:
    """Case: input_docx is good, but template path doesn't exist."""
    sample_d2p_cfg.template_pptx = "bad/path.pptx"
    with pytest.raises(FileNotFoundError, match="not found"):
        sample_d2p_cfg.validate_docx2pptx_pipeline_requirements()


def test_validate_docx2pptx_pipeline_requirements_raises_if_template_is_dir(
    sample_d2p_cfg: UserConfig, tmp_path: Path
) -> None:
    """Case: Input is good, and template exists, but it's a directory."""
    sample_d2p_cfg.template_pptx = str(tmp_path)
    with pytest.raises(ValueError, match="must be a file, not a folder"):
        sample_d2p_cfg.validate_docx2pptx_pipeline_requirements()


def test_validate_docx2pptx_pipeline_requirements_raises_if_output_is_file(
    sample_d2p_cfg: UserConfig, path_to_empty_docx: Path
) -> None:
    """Case: Output folder is a file instead of a directory"""
    sample_d2p_cfg.output_folder = str(path_to_empty_docx)
    with pytest.raises(ValueError, match="exists but is not a directory"):
        sample_d2p_cfg.validate_docx2pptx_pipeline_requirements()


# endregion

# region test_validate_pptx2docx_pipeline_requirements


def test_validate_pptx2docx_pipeline_requirements_valid_config(
    sample_p2d_cfg: UserConfig,
) -> None:
    """Case: baseline happy path works for validate_pptx2docx_pipeline_requirements"""
    sample_p2d_cfg.validate_pptx2docx_pipeline_requirements()


@pytest.mark.parametrize(
    argnames="config_input,expected_exception,match_pattern",
    argvalues=[
        # Case: Input path is None
        (UserConfig(input_pptx=None), ValueError, "No input pptx file specified"),
        # Case: Input pptx path does not exist on disk
        (
            UserConfig(input_pptx="this_file_doesn't_exist.pptx"),
            FileNotFoundError,
            "file not found",
        ),
    ],
)
def test_validate_pptx2docx_pipeline_requirements_raise_if_input_None_or_not_exist(
    config_input: UserConfig,
    expected_exception: type[Exception],
    match_pattern: str,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Cases: input_pptx is set to None, input_pptx is set to a path that doesn't exist."""
    with pytest.raises(expected_exception, match=match_pattern):
        config_input.validate_pptx2docx_pipeline_requirements()
    assert match_pattern in caplog.text


def test_validate_pptx2docx_pipeline_requirements_raises_if_input_pptx_is_dir(
    tmp_path: Path,
) -> None:
    """Case: input_pptx is set to a (valid) directory, not a file."""
    test_cfg = UserConfig(input_pptx=str(tmp_path))
    with pytest.raises(expected_exception=ValueError, match="not a file"):
        test_cfg.validate_pptx2docx_pipeline_requirements()


def test_validate_pptx2docx_pipeline_requirements_raises_if_template_not_exist(
    sample_p2d_cfg: UserConfig,
) -> None:
    """Case: input_pptx is good, but template_docx file does not exist."""
    sample_p2d_cfg.template_docx = "bad/path.docx"
    with pytest.raises(FileNotFoundError, match="not found"):
        sample_p2d_cfg.validate_pptx2docx_pipeline_requirements()


def test_validate_pptx2docx_pipeline_requirements_raises_if_template_is_dir(
    sample_p2d_cfg: UserConfig, tmp_path: Path
) -> None:
    """Case: Input is good, and template exists, but it's a directory."""
    sample_p2d_cfg.template_docx = str(tmp_path)
    with pytest.raises(ValueError, match="must be a file, not a folder"):
        sample_p2d_cfg.validate_pptx2docx_pipeline_requirements()


# Output folder does not exist
def test_validate_pptx2docx_pipeline_requirements_raises_if_output_is_file(
    sample_p2d_cfg: UserConfig, path_to_empty_docx: Path
) -> None:
    """Case: Output folder is a file instead of a directory"""
    sample_p2d_cfg.output_folder = str(path_to_empty_docx)
    with pytest.raises(ValueError, match="exists but is not a directory"):
        sample_p2d_cfg.validate_pptx2docx_pipeline_requirements()


# endregion


# endregion
