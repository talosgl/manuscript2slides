# tests/test_smoke.py
"""Smoke tests to ensure basic functionality works."""

import logging
from pathlib import Path

import pytest

from manuscript2slides.internals.define_config import UserConfig
from manuscript2slides.orchestrator import run_roundtrip_test


def test_roundtrip_works(
    path_to_sample_docx_with_everything: Path,
    path_to_empty_pptx: Path,
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Run roundtrip as a quick smoke test for the pipeline, as I used to do manually, but
    use test data."""
    # Arrange a config to emulate how we run roundtrip from CLI and GUI,
    # but use the test data fixtures intead of relying on user directories
    test_cfg = UserConfig(
        input_docx=path_to_sample_docx_with_everything,
        template_pptx=path_to_empty_pptx,
        output_folder=tmp_path,
    ).enable_all_options()

    with caplog.at_level(logging.DEBUG):
        original_docx, intermediate_pptx, final_docx = run_roundtrip_test(test_cfg)
    assert "success" in caplog.text
    assert original_docx and original_docx.exists()
    assert intermediate_pptx and intermediate_pptx.exists()
    assert final_docx and final_docx.exists()


def test_config_creation() -> None:
    """Test that we can create a config object."""
    cfg = UserConfig()
    assert cfg is not None
    assert cfg.chunk_type is not None


def test_config_with_defaults() -> None:
    """Test that demo config works."""
    cfg = UserConfig().with_defaults()
    assert cfg.input_docx is not None
