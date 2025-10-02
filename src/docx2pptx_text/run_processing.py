"""TODO"""
from docx.text.paragraph import Paragraph as Paragraph_docx
from docx.text.run import Run as Run_docx
from pptx.text.text import _Paragraph as Paragraph_pptx, _Run as Run_pptx  # type: ignore
from src.docx2pptx_text.utils import debug_print, detect_field_code_hyperlinks
from src.docx2pptx_text.formatting import copy_run_formatting_docx2pptx
from docx import document
from src.docx2pptx_text.models import SlideNotes
from docx.opc import constants
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement as OxmlElement_docx
from src.docx2pptx_text.formatting import copy_run_formatting_pptx2docx, _apply_experimental_formatting_from_metadata
from src.docx2pptx_text.annotations.restore_from_slides import safely_extract_comment_data, safely_extract_experimental_formatting_data
from src.docx2pptx_text import config

# region docx2pptx
def process_docx_paragraph_inner_contents(
    paragraph: Paragraph_docx, pptx_paragraph: Paragraph_pptx
) -> list[dict]:
    """
    Iterate through a paragraph's runs and hyperlinks, in document order, and:
    - copy it into the slide, with formatting
    - capture any additional metadata (like experimental formatting) that we cannot apply directly to the copied-over
        pptx objects
    """
    items_processed = False
    experimental_formatting_metadata = []

    for item in paragraph.iter_inner_content():
        items_processed = True
        if isinstance(item, Run_docx):

            # If this Run has a field code for instrText and it begins with HYPERLINK, this is an old-style
            # word hyperlink, which we cannot handle the same way as normal docx hyperlinks. But we try to detect
            # when it happens and report it to the user.
            field_code_URL = detect_field_code_hyperlinks(item)
            if field_code_URL:
                item.text = f" [Link: {field_code_URL}] "

            process_docx_run(item, pptx_paragraph, experimental_formatting_metadata)

        # Run and Hyperlink objects are peers in docx, but Hyperlinks can contain lists of Runs.
        # We check the item.url field because that seems the most reliable way to see if this is a
        # basic run versus a Hyperlink containing its own nested runs.
        # https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.hyperlink.Hyperlink.url
        elif hasattr(item, "url"):
            # Process all runs within the Hyperlink
            for run in item.runs:
                process_docx_run(
                    run,
                    pptx_paragraph,
                    experimental_formatting_metadata,
                    item.url,
                )
        # elif hasattr(item, "fragment"):
        #   ...
        #   TODO, leafy: We need to handle document anchors differently from other hyperlinks.
        #   We need to 1) process the nested runs as if it were a Hyperlink object, and
        #   2) preserve the anchor somewhere, maybe the experimental formatting metadata.
        #   https://python-docx.readthedocs.io/en/latest/api/text.html#docx.text.hyperlink.Hyperlink.fragment

        else:
            debug_print(f"Unknown content type in paragraph: {type(item)}")

    # Fallback: if no content was processed but paragraph has text
    if not items_processed and paragraph.text:
        debug_print(
            f"Fallback: paragraph has text but no runs/hyperlinks: {paragraph.text[:50]}"
        )
        pptx_run = pptx_paragraph.add_run()
        pptx_run.text = paragraph.text

    return experimental_formatting_metadata



def process_docx_run(
    run: Run_docx,
    pptx_paragraph: Paragraph_pptx,
    experimental_formatting_metadata: list,
    hyperlink: str | None = None,
) -> Run_pptx:
    """Copy a run from the docx parent to the pptx paragraph, including copying its formatting."""
    # Handle formatting

    pptx_run = pptx_paragraph.add_run()
    copy_run_formatting_docx2pptx(run, pptx_run, experimental_formatting_metadata)

    if hyperlink:
        pptx_run_url = pptx_run.hyperlink
        pptx_run_url.address = hyperlink

    return pptx_run
