"""Tests for UserConfig class definition file and related items"""

import pytest
from manuscript2slides.internals.config.define_config import (
    ChunkType,
    UserConfig,
    PipelineDirection,
)
import sys


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
