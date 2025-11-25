"""Shared fixtures"""

# tests/conftest.py
import pytest
from pathlib import Path
import tempfile
import shutil
from manuscript2slides.internals.config.define_config import UserConfig, ChunkType


@pytest.fixture
def temp_output_dir(tmp_path: Path) -> Path:
    """Temporary directory for test output files"""
    output = tmp_path / "output"
    output.mkdir()
    return output


@pytest.fixture
def sample_docx_with_formatting() -> Path:
    """Path to test docx with various formatting examples"""
    path = Path("tests/data/test_formatting.docx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture
def sample_d2p_cfg(
    sample_docx_with_formatting: Path, temp_output_dir: Path
) -> UserConfig:
    """Sample config object for docx2pptx testing"""
    return UserConfig(
        input_docx=str(sample_docx_with_formatting),  # Use real test file
        output_folder=str(temp_output_dir),
        chunk_type=ChunkType.HEADING_FLAT,
        experimental_formatting_on=True,
    )


@pytest.fixture
def clean_debug_env(monkeypatch: pytest.MonkeyPatch) -> pytest.MonkeyPatch:
    """Ensure debug env var is not set before test.
    Used by at least test_utils + test_cli."""
    # Pytest will temporarily remove it from THIS test/caller's view of the environment
    monkeypatch.delenv("MANUSCRIPT2SLIDES_DEBUG", raising=False)
    return monkeypatch


# TODO: Assess if this is really needed because I was just guessing while learning
@pytest.fixture
def sample_config_toml() -> Path:
    """Path to test a config toml"""
    path = Path("tests/data/test_config.toml")
    assert path.exists(), f"Test file not found: {path}"
    return path
