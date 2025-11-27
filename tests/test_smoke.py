# tests/test_smoke.py
"""Smoke tests to ensure basic functionality works."""
import pytest
from manuscript2slides.internals.define_config import UserConfig


def test_config_creation() -> None:
    """Test that we can create a config object."""
    cfg = UserConfig()
    assert cfg is not None
    assert cfg.chunk_type is not None


def test_config_with_defaults() -> None:
    """Test that demo config works."""
    cfg = UserConfig().with_defaults()
    assert cfg.input_docx is not None
