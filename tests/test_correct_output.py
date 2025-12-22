"""Tests for correctness of pipeline output based on manual test cases performed during development."""

# pyright: reportAttributeAccessIssue=false
import pytest
from pathlib import Path

import pptx
from pptx import presentation
from pptx.slide import Slide
from pptx.text.text import TextFrame

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


def get_slide_notes_text(slide: Slide) -> TextFrame:
    """Helper to modularize getting a slide's notes frame for testing against."""
    assert (
        slide.has_notes_slide is True
        and slide.notes_slide is not None
        and slide.notes_slide.notes_text_frame is not None
    ), f"Test cannot proceed because we expected the slide, {slide.slide_id} to have a speaker notes section and it does not."

    notes_text_frame: TextFrame = slide.notes_slide.notes_text_frame

    return notes_text_frame


# region docx2pptx tests

""" 
When converting from sample_doc.docx -> standard pptx output, with these options:
    - preserve experimental formatting
    - keep all annotations
    - preserve metadata in speaker notes
...the results should match the assertions in the tests below.

# TODO: Consider if we should do similar testing for when options are disabled. 
# (E.g., experimental formatting on, but speaker notes are empty.)
"""


def test_where_are_data_slide(output_pptx: Path) -> None:
    """
    Find the slide with the 'Where are Data?' title text. Test slide's text formatting
    and the contents of slide notes against expectations.

    This verifies:
        - standard formatting (font color, italics) is copied from docx runs to pptx runs
        - standard formatting (bold) gets copied from docx paragraphs to pptx paragraphs
        - the above two things BOTH happen and don't stomp each other unexpectedly (we do expect run font to stomp paragraph font)
        - typeface gets copied from docx paragraphs OR runs to pptx paragraphs
        - when display_comments is enabled, they get copied into the speaker notes successfully
        - when preserve_docx_metadata_in_speaker_notes is enabled, the Heading's font data is copied into the speaker notes successfully
    """

    # Action: note that the "Action" here is running the docx2pptx pipeline, which we do
    # in a fixture one time instead of repeatedly for many tests.
    prs = pptx.Presentation(output_pptx)

    # Arrange: Find the slide and paragraph containing the formatted text we are going to assert against.
    slide_result = find_first_slide_containing(prs, "Where are Data?")

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    # Assert
    # === Test case: The paragraph should be bolded at the paragraph-level.
    assert para.font.bold is True
    assert para.font.name == "Times New Roman"

    # === Test case: "are" within "Where are Data?" should be in red. ===
    # Arrange: Find the run within the slide we expect to have red color.
    # We have to do exact match, or the URL run will be returned.
    run_result = find_run_in_para_with_exact_match(para, "are")

    assert (
        run_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    assert run_result.font.color is not None
    assert (
        hasattr(run_result.font.color, "rgb") and run_result.font.color.rgb is not None
    )
    assert run_result.font.color.rgb == RGBColor_pptx(0xFF, 0x00, 0x00)  # red
    assert run_result.font.italic is True

    # === Test Case: Paragraph text should contain link text from a field code hyperlink.
    assert "http" in para.text

    notes_text_frame: TextFrame = get_slide_notes_text(slide)

    # === Test Case: The speaker notes should contain copied comments, because UserConfig.display_comments is true in conftest.output_pptx's pipeline run.
    assert "COMMENTS FROM SOURCE DOCUMENT" in notes_text_frame.text
    assert '"text": "What happens if' in notes_text_frame.text
    assert "What happens if there's a threaded comment?" in notes_text_frame.text

    # === Test Case: The speaker notes should contain formatting data for the heading on this slide, because UserConfig.preserve_docx_metadata_in_speaker_notes was true
    assert '"headings": [' in notes_text_frame.text
    assert '"name": "Heading 1"' in notes_text_frame.text


def test_author_slide(output_pptx: Path) -> None:
    """Find slide containing 'by J. King-Yost' and verify:
    - In paragraphs where there is only heading-level font data in the docx, it is copied to the paragraph-level
    in the pptx successfully, including color, italics, and typeface
    - When preserve_docx_metadata_in_speaker_notes is enabled, the Heading's font data is copied into the speaker notes successfully
    """

    prs = pptx.Presentation(output_pptx)
    slide_result = find_first_slide_containing(prs, "by J. King-Yost")

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    # === Test Case: the text should be gray, in italics, and should be in Times New Roman typeface
    # NOTE: Because this is a heading, the formatting is applied at the paragraph-level, not the run-level.
    # We will explicitly check the paragraph's font here; it is possible the test may need to be adapted
    # to be more flexible and to check both (either/or) in the future.

    assert para.font.color is not None
    assert hasattr(para.font.color, "rgb") and para.font.color.rgb is not None
    assert para.font.color.rgb == RGBColor_pptx(0x99, 0x99, 0x99)  # gray
    assert para.font.italic is True
    assert para.font.name == "Times New Roman"

    notes_text_frame: TextFrame = get_slide_notes_text(slide)

    assert "JSON METADATA FROM SOURCE DOCUMENT" in notes_text_frame.text
    assert '"headings": [' in notes_text_frame.text
    assert '"name": "Heading 3"' in notes_text_frame.text


# TODO:
"""
In the slide containing a paragraph beginning with: "In a cold concrete underground tunnel"
- "In a cold concrete underground tunnel" should be a functional hyperlink to https://dataepics.webflow.io/stories/where-are-data
- "splayed" should be bolded
- "three dozen directions" should be in italics
- "buzzing" should be highlighted in yellow
- "Dust covers the cables" should be in red text
- "She could wipe them clean without too much risk" should be underlined
- {!} "read those stories" should be double-underlined - current behavior output is single-underline. I can't remember if this was a limitation we accepted or not: investigate later.
- Speaker notes should contain:
	- "COMMENTS FROM SOURCE DOCUMENT:" and "Sample comment"
	- "FOOTNOTES FROM SOURCE DOCUMENT:" and "1. James Griffiths."
	- "START OF JSON METADATA FROM SOURCE DOCUMENT" and "experimental_formatting" and YELLOW and Sample Comment and James Griffiths...
    """

# TODO:
"""
The slide containing "Vedantam, Shankar" in the main body of the slide, probably the last slide:
- speaker notes on this slide should contain "ENDNOTES FROM SOURCE DOCUMENT" and "sample endnote"
- "START OF JSON METADATA FROM SOURCE DOCUMENT" and "endnotes" and "reference_text"
"""
# endregion
