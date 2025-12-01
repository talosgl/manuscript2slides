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


# region test validate()
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
        # Output folder cannot be empty string
        (
            UserConfig(input_docx="input.docx", output_folder=""),
            ValueError,
            "cannot be empty string; use None for default",
        ),
        # Range checks
        (
            UserConfig(range_start=5, range_end=2),
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


# region for later
# TODO: a bunch more tests in here

# TODO: test from_toml

# TODO: test direction property...?

# TODO: might be low value-to-effort, but, test all the get_*_path() ?

# TODO: test enable_all_options

# TODO: test save_toml (?)

# TODO: test config_to_dict
# endregion
