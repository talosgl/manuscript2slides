"""Shared test helper functions."""

# mypy: disable-error-code="import-untyped"
# pyright: reportArgumentType=false

from pptx import presentation
from pptx.slide import Slide

from pptx.text.text import _Paragraph as Paragraph_pptx
from pptx.text.text import _Run as Run_pptx
from manuscript2slides.processing.populate_docx import get_slide_paragraphs


def find_first_slide_containing(
    prs: presentation.Presentation, search_text: str
) -> tuple[Slide, Paragraph_pptx] | None:
    """Find the first slide that contains the given text anywhere in its shapes (but not searching its speaker notes).
    Returns None if no slide found with this text."""

    if not prs.slides:
        return None

    slides = list(prs.slides)

    for slide in slides:
        paragraphs: list[Paragraph_pptx] = get_slide_paragraphs(slide)

        for para in paragraphs:
            if search_text in para.text:
                return (slide, para)

    return None


def find_first_run_in_para_containing(
    para: Paragraph_pptx, search_text: str
) -> Run_pptx | None:
    """Given a paragraph from a slide, find the run containing the search text, or None if not found."""
    if not para.text or (len(para.runs) < 1):
        return None

    for run in para.runs:
        if search_text in run.text:
            return run

    return None


def find_run_in_para_with_exact_match(
    para: Paragraph_pptx, exact_text: str
) -> Run_pptx | None:
    """Find a run whose text exactly matches the given text (not substring match)."""
    if not para.text or len(para.runs) < 1:
        return None

    for run in para.runs:
        if run.text == exact_text:
            return run

    return None


def get_speaker_notes_text(slide: Slide) -> str | None:
    """Get the text content of a slide's speaker notes, or empty string if no notes."""
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
        return slide.notes_slide.notes_text_frame.text
    else:
        return None
