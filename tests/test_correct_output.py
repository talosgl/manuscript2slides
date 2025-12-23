"""Tests for correctness of pipeline output based on manual test cases performed during development."""

# pyright: reportAttributeAccessIssue=false
import pytest
from pathlib import Path

import pptx
from pptx import presentation
from pptx.slide import Slide
from pptx.text.text import TextFrame
from pptx.text.text import _Hyperlink as Hyperlink_pptx
from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE as MSO_TEXT_UNDERLINE_TYPE_pptx

from pptx.text.text import _Paragraph as Paragraph_pptx
from pptx.text.text import _Run as Run_pptx
from pptx.dml.color import RGBColor as RGBColor_pptx
from docx.text.hyperlink import Hyperlink as Hyperlink_docx


from manuscript2slides.internals.define_config import UserConfig
from tests import helpers


# region docx2pptx tests

""" 
When converting from sample_doc.docx -> standard pptx output, with these options:
    - preserve experimental formatting
    - keep all annotations
    - preserve metadata in speaker notes
...the results should match the assertions in the tests below.

# TODO: Consider if we should add similar tests for when options are disabled. 
# (E.g., experimental formatting on, but preserve_docx_metadata_in_speaker_notes is False)
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
    slide_result = helpers.find_first_slide_containing(prs, "Where are Data?")

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    # Assert
    # Case: The paragraph should be bolded at the paragraph-level
    assert para.font.bold is True

    # Case: "are" within "Where are Data?" should be in red
    # We have to do exact match, or the URL run will be returned
    run_result = helpers.find_first_pptx_run_in_para_with_exact_match(para, "are")

    # Verify color object exists and has RGB value before checking specific color
    assert run_result.font.color is not None
    assert (
        hasattr(run_result.font.color, "rgb") and run_result.font.color.rgb is not None
    )
    assert run_result.font.color.rgb == RGBColor_pptx(0xFF, 0x00, 0x00)  # red
    assert run_result.font.italic is True

    # Case: Paragraph text should contain link text from a field code hyperlink
    assert "http" in para.text

    notes_text_frame: TextFrame = helpers.get_slide_notes_text(slide)

    # Case: Speaker notes should contain copied comments (UserConfig.display_comments is true)
    assert "COMMENTS FROM SOURCE DOCUMENT" in notes_text_frame.text
    assert '"text": "What happens if' in notes_text_frame.text
    assert "What happens if there's a threaded comment?" in notes_text_frame.text

    # Case: Speaker notes should contain formatting data for the heading (preserve_docx_metadata_in_speaker_notes is true)
    assert '"headings": [' in notes_text_frame.text
    assert '"name": "Heading 1"' in notes_text_frame.text


def test_author_slide(output_pptx: Path) -> None:
    """Find slide containing 'by J. King-Yost' and verify:
    - In paragraphs where there is only heading-level font data in the docx, it is copied to the paragraph-level
    in the pptx successfully, including color, italics, and typeface
    - When preserve_docx_metadata_in_speaker_notes is enabled, the Heading's font data is copied into the speaker notes successfully
    """

    prs = pptx.Presentation(output_pptx)
    slide_result = helpers.find_first_slide_containing(prs, "by J. King-Yost")

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    # Case: The text should be gray, in italics, and in Times New Roman typeface
    # NOTE: Because this is a heading, the formatting is applied at the paragraph-level, not the run-level.
    # We will explicitly check the paragraph's font here; it is possible the test may need to be adapted
    # to be more flexible and to check both (either/or) in the future.

    # Verify color object exists and has RGB value before checking specific color
    assert para.font.color is not None
    assert hasattr(para.font.color, "rgb") and para.font.color.rgb is not None
    assert para.font.color.rgb == RGBColor_pptx(0x99, 0x99, 0x99)  # gray
    assert para.font.italic is True

    notes_text_frame: TextFrame = helpers.get_slide_notes_text(slide)

    assert "JSON METADATA FROM SOURCE DOCUMENT" in notes_text_frame.text
    assert '"headings": [' in notes_text_frame.text
    assert '"name": "Heading 3"' in notes_text_frame.text


def test_in_a_cold_concrete_underground_tunnel_slide(output_pptx: Path) -> None:
    """
    Test the following cases:
    Slide body:
    - standard formatting is preserved via italic, bold, color, and underline runs individually

    - experimental formatting via highlight text is preserved
    Speaker notes:
    - footnotes are copied when display_footnotes is enabled, including any hyperlinks
    - comments are copied when display_comments is enabled
    - experimental formatting data is copied when preserve_metadata... is enabled
    - footnotes data is copied when metadata preservation is enabled, including its hyperlinks
    """
    prs = pptx.Presentation(output_pptx)
    slide_result = helpers.find_first_slide_containing(
        prs, "In a cold concrete underground tunnel"
    )

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    # Case: Hyperlinks are preserved
    # "In a cold concrete underground tunnel" should be a functional hyperlink to https://dataepics.webflow.io/stories/where-are-data
    hyperlink_run = helpers.find_first_pptx_run_in_para_containing(
        para, "In a cold concrete underground"
    )

    # Verify hyperlink object exists and is correct type before checking URL
    assert (
        hyperlink_run.hyperlink is not None
        and isinstance(hyperlink_run.hyperlink, Hyperlink_pptx)
        and hyperlink_run.hyperlink.address is not None
    )
    assert (
        hyperlink_run.hyperlink.address
        == "https://dataepics.webflow.io/stories/where-are-data"
    )

    # Case: Hyperlinks AND standard formatting is preserved in the same run
    # the "tunnel" part of the clause should be in italics and maintain the same hyperlink as the above.
    hyperlink_run_formatted = helpers.find_first_pptx_run_in_para_with_exact_match(
        para, "tunnel"
    )  # Tunnel shows up in multiple spots in this paragraph so we're relying on this being the first one

    # Verify hyperlink object exists and is correct type before checking URL
    assert (
        hyperlink_run_formatted.hyperlink is not None
        and isinstance(hyperlink_run_formatted.hyperlink, Hyperlink_pptx)
        and hyperlink_run_formatted.hyperlink.address is not None
    )
    assert (
        hyperlink_run_formatted.hyperlink.address
        == "https://dataepics.webflow.io/stories/where-are-data"
    )
    assert hyperlink_run_formatted.font.italic is True

    # Case: Standard formatting is preserved - "splayed" should be bolded
    bolded_run = helpers.find_first_pptx_run_in_para_with_exact_match(para, "splayed")
    assert bolded_run.font.bold is True

    # Case: "three dozen directions" should be italic
    italic_run = helpers.find_first_pptx_run_in_para_with_exact_match(
        para, "three dozen directions"
    )
    assert italic_run.font.italic is True

    # Case: "Dust covers the cables" should be in red text
    red_run = helpers.find_first_pptx_run_in_para_containing(
        para, "Dust covers the cables"
    )
    # Verify color object exists and has RGB value before checking specific color
    assert red_run.font.color is not None
    assert hasattr(red_run.font.color, "rgb") and red_run.font.color.rgb is not None
    assert red_run.font.color.rgb == RGBColor_pptx(0xFF, 0x00, 0x00)  # red

    # Case: "She could wipe them clean without too much risk" should be underlined
    ul_run = helpers.find_first_pptx_run_in_para_containing(
        para, "She could wipe them clean without too much risk"
    )
    assert ul_run.font.underline is True

    # Case: Experimental formatting is preserved ("buzzing" should be highlighted in yellow)
    highlight_run = helpers.find_first_pptx_run_in_para_with_exact_match(
        para, "buzzing"
    )
    assert helpers.run_has_highlight(highlight_run)
    hl_color = helpers.get_run_highlight_color(highlight_run)
    assert hl_color == "FFFF00"

    # Case: "read those stories" should be underlined... double-underlined!
    double_ul_run = helpers.find_first_pptx_run_in_para_containing(
        para, "read those stories"
    )
    assert double_ul_run.font.underline == MSO_TEXT_UNDERLINE_TYPE_pptx.DOUBLE_LINE

    # Speaker notes checks:
    notes_text_frame = helpers.get_slide_notes_text(slide)

    # Case: Comment is copied in when display_comments is enabled
    # Verify both comment section header and content are present
    assert (
        "COMMENTS FROM SOURCE DOCUMENT" in notes_text_frame.text
        and "Sample comment" in notes_text_frame.text
    )

    # Case: Footnote is copied in when display_footnotes is enabled
    # Verify footnote section, content, and hyperlink are all present
    assert (
        "FOOTNOTES FROM SOURCE DOCUMENT" in notes_text_frame.text
        and "1. James Griffiths." in notes_text_frame.text
        and "Hyperlinks:" in notes_text_frame.text
        and r"https://www.cnn.com/2019/07/25/asia/internet-undersea-cables-intl-hnk/index.html"
        in notes_text_frame.text
    )

    # Case: Footnote JSON is preserved when "preserve_metadata..." is enabled, with its hyperlink
    assert '"hyperlinks": [' in notes_text_frame.text

    # Case: experimental formatting data is copied in when "preserve_metadata..." is enabled
    assert "JSON METADATA FROM SOURCE DOCUMENT" in notes_text_frame.text
    # Verify the experimental_formatting JSON structure and all expected fields for "buzzing" highlight
    assert (
        '"experimental_formatting": [' in notes_text_frame.text
        and '"ref_text": "buzzing"' in notes_text_frame.text
        and '"highlight_color_enum": "YELLOW"' in notes_text_frame.text
        and '"formatting_type": "highlight"' in notes_text_frame.text
    )
    # Verify the footnote data was also preserved in the JSON metadata
    assert (
        '"footnotes": [' in notes_text_frame.text
        and '"hyperlinks": [' in notes_text_frame.text
        and '"note_type": "footnote"' in notes_text_frame.text
    )


def test_endnotes_slide(output_pptx: Path) -> None:
    """
    Find the slide containing "Vedantam, Shankar" in the main body of the slide, probably the last slide, and
    test that the endnote attached to this gets added to the speaker notes.
    - speaker notes on this slide should contain "ENDNOTES FROM SOURCE DOCUMENT" and "sample endnote"
    - "START OF JSON METADATA FROM SOURCE DOCUMENT" and "endnotes" and "reference_text"
    """
    prs = pptx.Presentation(output_pptx)
    slide_result = helpers.find_first_slide_containing(prs, "Vedantam, Shankar")

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    notes_text_frame = helpers.get_slide_notes_text(slide)

    # Case: Endnotes get copied into the speaker notes when display_endnotes is True
    assert "ENDNOTES FROM SOURCE DOCUMENT" in notes_text_frame.text
    assert "Another sample endnote with a url" in notes_text_frame.text
    assert "Hyperlinks:" in notes_text_frame.text
    assert (
        r"https://dataepics.webflow.io/stories/where-are-data" in notes_text_frame.text
    )

    # Case: Endnotes metadata get copied into the speaker notes when preserve metadata is True
    assert (
        "JSON METADATA FROM SOURCE DOCUMENT" in notes_text_frame.text
        and '"endnotes": [' in notes_text_frame.text
        and '"hyperlinks": [' in notes_text_frame.text
        and '"reference_text": "Vedantam, Shankar, and Iain McGilchrist.'
        in notes_text_frame.text
        and '"note_type": "endnote"' in notes_text_frame.text
    )


def test_vanilla_docx_theme_font_does_not_carry_into_pptx_output(
    output_pptx: Path,
) -> None:
    """
    Test if docx theme font is copied into pptx. It shouldn't be set at the paragraph or run level because
    the pptx's theme's fonts should be respected unless a run explicitly has a font typeface set.
    """
    prs = pptx.Presentation(output_pptx)

    slide_result = helpers.find_first_slide_containing(
        prs, "In a cold concrete underground tunnel"
    )

    assert (
        slide_result is not None
    ), f"Test cannot proceed because the required text could not be found."

    slide, para = slide_result

    assert para.runs[0].font.name is None


# endregion


# region pptx2docx
# TODO: Need to add equivalent tests for the reverse pipeline.

# endregion
