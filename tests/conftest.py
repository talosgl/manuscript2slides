# tests/conftest.py
import pytest
from pathlib import Path
import tempfile
import shutil
from manuscript2slides.internals.config.define_config import UserConfig, ChunkType


@pytest.fixture
def temp_dir():
    """Create a temp directory for test files."""

    temp = Path(tempfile.mkdtemp())
    yield temp
    shutil.rmtree(temp)  # cleanup after test


@pytest.fixture
def sample_docx_with_formatting():
    """Provide the path to a sample test document"""
    return Path("tests/data/test_formatting.docx")


def sample_d2p_cfg():
    return UserConfig(
        input_docx="/data/test_formatting.docx",
        chunk_type=ChunkType.HEADING_FLAT,
        experimental_formatting_on=True,
    )
