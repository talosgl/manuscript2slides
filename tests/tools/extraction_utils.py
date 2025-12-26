"""
Shared utilities for test fixture extraction scripts.

This module provides common helper functions used across the docx and pptx
extraction tools to reduce code duplication.
"""

from pprint import pprint
import sys

from docx.shared import RGBColor as RGBColor_docx
from pptx.dml.color import RGBColor as RGBColor_pptx


# XML namespace constants for Office documents
XML_NAMESPACE_WORD: str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS: dict[str, str] = {"w": XML_NAMESPACE_WORD}


def rgb_to_hex(rgb_color: RGBColor_docx | RGBColor_pptx) -> str:
    """
    Convert RGBColor object to hex string format.

    Args:
        rgb_color: RGBColor object (acts like a tuple with [0], [1], [2] indexing)

    Returns:
        Hex color string (e.g., "FF0000" for red)
    """
    rgb_int = (rgb_color[0] << 16) | (rgb_color[1] << 8) | rgb_color[2]
    return f"{rgb_int:06X}"


def safe_pprint(data: list[dict], width: int = 120, item_type: str = "items") -> None:
    """
    Pretty print data with graceful handling of Unicode encoding errors.

    Args:
        data: Data structure to print
        width: Maximum line width for pprint
        item_type: Type description for fallback message (e.g., "slides", "paragraphs")
    """
    print("\n" + "=" * 80)
    print("EXTRACTED DATA:")
    print("=" * 80)
    try:
        pprint(data, width=width)
    except UnicodeEncodeError:
        print(
            f"(Console encoding issue - data contains special characters)",
            file=sys.stderr,
        )
        print(f"Extracted {len(data)} {item_type}")


def filter_none_keep_false(data_dict: dict) -> dict:
    """
    Filter out None values from a dictionary but keep False values.

    This is useful for formatting data where False is meaningful (e.g., bold=False)
    but None indicates the property wasn't set.

    Args:
        data_dict: Dictionary to filter

    Returns:
        New dictionary with None values removed
    """
    return {k: v for k, v in data_dict.items() if v is not None}
