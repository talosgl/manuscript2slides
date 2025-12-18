"""Tests for correctness of pipeline output based on manual test cases."""

# pyright: reportAttributeAccessIssue=false
import pytest
from pathlib import Path

import pptx
from pptx import presentation
from pptx.slide import Slide

from pptx.text.text import _Paragraph as Paragraph_pptx
from pptx.text.text import _Run as Run_pptx
from pptx.dml.color import RGBColor as RGBColor_pptx


from manuscript2slides.internals.define_config import UserConfig
from tests.helpers import (
    find_first_slide_containing,
    find_first_run_in_para_containing,
    get_speaker_notes_text,
    find_run_in_para_with_exact_match,
)


# region docx2pptx tests
def test_where_are_data_slide(output_pptx: Path) -> None:
    prs = pptx.Presentation(output_pptx)
    sld_result = find_first_slide_containing(prs, "Where are Data?")

    assert (
        sld_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    _, para = sld_result

    run_result = find_run_in_para_with_exact_match(para, "are")

    assert (
        run_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    assert run_result.font.color is not None
    assert (
        hasattr(run_result.font.color, "rgb") and run_result.font.color.rgb is not None
    )
    assert run_result.font.color.rgb == RGBColor_pptx(0xFF, 0x00, 0x00)
    assert run_result.font.italic is True


# endregion
