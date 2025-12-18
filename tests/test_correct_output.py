"""Tests for correctness of pipeline output based on manual test cases performed during development."""

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

""" 
When converting from sample_doc.docx -> standard pptx output, with these options:
    - preserve experimental formatting
    - keep all annotations
    - preserve metadata in speaker notes
...the results should match the assertions in the below tests.

# TODO: Consider if we should do similar testing for when options are disabled. 
# (E.g., experimental formatting on, but speaker notes are empty.)
"""


def test_where_are_data_slide(output_pptx: Path) -> None:
    """Find the slide with the 'Where are Data?' title text. Test slide's text formatting
    and the contents of slide notes against expectations."""
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

    # TODO add the other test cases for this slide:
    """
    - paragraph text should contain link text from a field code hyperlink, maybe look for https in the string
	- the speaker notes should contain 
        - "COMMENTS FROM SOURCE DOCUMENT" and
        - "What happens if there's a threaded comment?"
        - AND "heading" data
    """


# TODO:
"""

- The slide containing "by J. King-Yost":
	- the text should be gray and in italics
	- the speaker notes should contain:
```
START OF JSON METADATA FROM SOURCE DOCUMENT  
========================================

{  
 "headings": [  
 {  
 "text": "by J. King-Yost",  
 "style_id": "Heading3",  
 "name": "Heading 3"  
 }  
 ]  
}

========================================  
END OF JSON METADATA FROM SOURCE DOCUMENT
```
"""

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
