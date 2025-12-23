"""Shared fixtures"""

# tests/conftest.py
import pytest
from pathlib import Path
from manuscript2slides.internals.define_config import UserConfig, ChunkType

from manuscript2slides.orchestrator import run_pipeline


@pytest.fixture
def temp_output_dir(tmp_path: Path) -> Path:
    """Temporary directory for test output files"""
    output = tmp_path / "output"
    output.mkdir()
    return output


@pytest.fixture
def path_to_sample_docx_with_formatting() -> Path:
    """Path to test docx with various formatting examples"""
    path = Path("tests/data/test_formatting.docx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture
def path_to_sample_pptx_with_formatting() -> Path:
    """Path to test pptx with various formatting examples"""
    path = Path("tests/data/test_formatting.pptx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture(scope="session")
def path_to_empty_pptx() -> Path:
    """Path to an empty slide deck (no slides) object for the purpose of instantiating
    new decks with pptx.Presentation() constructor."""
    path = Path("tests/data/pptx_template.pptx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture
def path_to_pptx_w_twenty_empty_slides() -> Path:
    """Path to a slide deck object with 20 empty slides for the purpose of instantiating
    new decks with pptx.Presentation() constructor, and populating slides with test data.
    """
    path = Path("tests/data/pptx_20_empty_slides.pptx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture
def path_to_empty_docx() -> Path:
    """Path to an empty word docx file. Currently unused"""
    path = Path("tests/data/docx_template.docx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture(scope="session")
def path_to_sample_docx_with_everything() -> Path:
    """Path to a copy of the standard sample_doc.docx that lives in tests/data."""
    path = Path("tests/data/sample_doc.docx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture(scope="session")
def path_to_sample_pptx_with_everything() -> Path:
    """Path to a pptx in tests/data that used a custom template during docx2pptx run, for use
    in reverse pipeline tests."""
    path = Path("tests/data/custom_template_output.docx")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture
def sample_d2p_cfg(
    path_to_sample_docx_with_formatting: Path,
    temp_output_dir: Path,
    path_to_empty_pptx: Path,
) -> UserConfig:
    """Sample config object for docx2pptx testing"""
    return UserConfig(
        input_docx=path_to_sample_docx_with_formatting,  # Use real test file
        template_pptx=path_to_empty_pptx,
        output_folder=temp_output_dir,
        chunk_type=ChunkType.HEADING_FLAT,
        experimental_formatting_on=True,
    )


@pytest.fixture
def sample_p2d_cfg(
    path_to_sample_pptx_with_formatting: Path,
    temp_output_dir: Path,
    path_to_empty_docx: Path,
) -> UserConfig:
    """Sample config object for pptx2docx testing"""
    return UserConfig(
        input_pptx=path_to_sample_pptx_with_formatting,  # Use real test file
        template_docx=path_to_empty_docx,
        output_folder=temp_output_dir,
    )


@pytest.fixture
def clean_debug_env(monkeypatch: pytest.MonkeyPatch) -> pytest.MonkeyPatch:
    """Ensure debug env var is not set before test.
    Used by at least test_utils + test_cli."""
    # Pytest will temporarily remove it from THIS test/caller's view of the environment
    monkeypatch.delenv("MANUSCRIPT2SLIDES_DEBUG", raising=False)
    return monkeypatch


@pytest.fixture
def sample_config_toml() -> Path:
    """Path to test a config toml"""
    path = Path("tests/data/test_config.toml")
    assert path.exists(), f"Test file not found: {path}"
    return path


@pytest.fixture(scope="session")
def session_temp_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Session-scoped temporary directory that persists across all tests.

    Uses tmp_path_factory instead of tmp_path because tmp_path is function-scoped
    and cannot be used as a dependency for session-scoped fixtures.
    """
    # pytest handles cleanup automatically for tmp_path_factory
    return tmp_path_factory.mktemp("session")


@pytest.fixture(scope="session")
def output_pptx(
    path_to_sample_docx_with_everything: Path,
    path_to_empty_pptx: Path,
    session_temp_dir: Path,
) -> Path:
    """Run the pipeline once for the entire test session with every option enabled."""
    cfg = UserConfig(
        input_docx=path_to_sample_docx_with_everything,
        template_pptx=path_to_empty_pptx,
        output_folder=session_temp_dir,
    ).enable_all_options()

    output_filepath = run_pipeline(cfg)
    return output_filepath


@pytest.fixture(scope="session")
def output_pptx_default_options(
    path_to_sample_docx_with_everything: Path,
    path_to_empty_pptx: Path,
    session_temp_dir: Path,
) -> Path:
    """Run the pipeline once for the entire test session with only default options set.

    Defaults (as of 2025-12-23):
    display_comments = False
    display_endnotes = False
    display_footnotes = False
    experimental_formatting_on = True
    preserve_docx_metadata_in_speaker_notes = False
    chunk_type = PARAGRAPH
    comments_keep_author_and_date = True
    comments_sort_by_date = True
    direction = (determined by input) DOCX_TO_PPTX

    """
    cfg = UserConfig(
        input_docx=path_to_sample_docx_with_everything,
        template_pptx=path_to_empty_pptx,
        output_folder=session_temp_dir,
    )

    output_filepath = run_pipeline(cfg)
    return output_filepath


@pytest.fixture(scope="session")
def output_docx(
    path_to_sample_pptx_with_everything: Path,
    path_to_empty_docx: Path,
    session_temp_dir: Path,
) -> Path:
    """Run the pipeline once for the entire test session."""
    cfg = UserConfig(
        input_pptx=path_to_sample_pptx_with_everything,
        template_docx=path_to_empty_docx,
        output_folder=session_temp_dir,
    )

    output_filepath = run_pipeline(cfg)
    return output_filepath
