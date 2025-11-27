"""Wide scoped tests that touch many parts of the pipeline and program to ensure no catastrophic failures during golden path runs."""

# tests/test_integration.py
import pytest
from pathlib import Path
from manuscript2slides.internals.define_config import UserConfig
from manuscript2slides.orchestrator import run_roundtrip_test
from docx import Document


def test_round_trip_preserves_basic_plaintext_content(
    path_to_sample_docx_with_everything: Path,
) -> None:
    """Integration test: docx -> pptx -> docx preserves content"""

    # Arrange:
    # Get the original docx's text to test against later
    original_docx = Document(str(path_to_sample_docx_with_everything))
    original_text = "\n".join(
        p.text for p in original_docx.paragraphs if p.text.strip()
    )

    # Make a config to send to round_trip
    test_cfg = UserConfig.with_defaults().enable_all_options()
    test_cfg.input_docx = str(path_to_sample_docx_with_everything)

    _, _, final_docx = run_roundtrip_test(test_cfg)

    final_doc = Document(str(final_docx))
    final_text = "\n".join(p.text for p in final_doc.paragraphs if p.text.strip())

    # Basic check: did we keep most of the text?
    assert len(final_text) > (len(original_text) * 0.8), "Lost too much content"


def test_docx2pptx_creates_valid_output(
    path_to_sample_docx_with_formatting: Path, tmp_path: Path
) -> None:
    """Integration: docx → pptx works"""
    # TODO


def test_pptx2docx_creates_valid_output(tmp_path: Path) -> None:
    """Integration: pptx → docx works"""
    # TODO



# def test_pipeline_fails_gracefully_on_missing_input():
#     """Error path: missing input file raises, doesn't crash"""
#     cfg = UserConfig.with_defaults()
#     cfg.input_docx = "nonexistent.docx"
    
#     with pytest.raises(FileNotFoundError):
#         run_pipeline(cfg)
#     # Could add: check logs, check no partial files created, etc.


# def test_pipeline_fails_gracefully_on_invalid_config():
#     """Error path: invalid config caught before pipeline runs"""
#     cfg = UserConfig()
#     cfg.input_docx = None  # Invalid - no input
    
#     with pytest.raises(ValueError):
#         run_pipeline(cfg)