# io.py
"""File I/O operations for docx and pptx files."""

import logging
from datetime import datetime
from pathlib import Path
from typing import TypeVar

import docx
import pptx
from docx import document
from docx.text.paragraph import Paragraph as Paragraph_docx
from pptx import presentation
from pptx.slide import Slide

from manuscript2slides.internals import constants
from manuscript2slides.internals.config.define_config import UserConfig
from manuscript2slides.processing.populate_docx import get_slide_paragraphs
from manuscript2slides.internals.run_context import get_pipeline_run_id

log = logging.getLogger("manuscript2slides")

OUTPUT_TYPE = TypeVar("OUTPUT_TYPE", document.Document, presentation.Presentation)


# region Path Helpers
def validate_path(user_path: str | Path) -> Path:
    """Ensure filepath exists and is a file."""
    path = Path(user_path)
    pipeline_id = get_pipeline_run_id()
    if not path.exists():
        log.error(f"File not found: {user_path} [pipeline:{pipeline_id}]")
        raise FileNotFoundError(f"File not found: {user_path}")
    if not path.is_file():
        log.error(
            f"Path is not a file (might be a directory): {user_path} [pipeline:{pipeline_id}]"
        )
        raise ValueError(f"Path is not a file: {user_path}")
    return path


def validate_pptx_path(user_path: str | Path) -> Path:
    """Validates the pptx template filepath exists and is actually a pptx file."""
    path = validate_path(user_path)

    # Get pipeline ID for logs.
    pipeline_id = get_pipeline_run_id()

    # Verify it's the right extension:
    if path.suffix.lower() == ".ppt":
        log.error(f"Unsupported .ppt file: {path} [pipeline:{pipeline_id}]")
        raise ValueError(
            "This tool only supports .pptx files right now. Please convert your .ppt file to .pptx format first."
        )
    if path.suffix.lower() != ".pptx":
        log.error(
            f"Wrong file extension: expected .pptx, got {path.suffix} [pipeline:{pipeline_id}]"
        )
        raise ValueError(f"Expected a .pptx file, but got: {path.suffix}")
    return path


def validate_docx_path(user_path: str | Path) -> Path:
    """Validates the user-provided filepath exists and is actually a docx file."""
    path = validate_path(user_path)

    # Get pipeline ID for logs.
    pipeline_id = get_pipeline_run_id()

    # Verify it's the right extension:
    if path.suffix.lower() == ".doc":
        log.error(f"Unsupported .doc file: {path} [pipeline:{pipeline_id}]")
        raise ValueError(
            "This tool only supports .docx files right now. Please convert your .doc file to .docx format first."
        )
    if path.suffix.lower() != ".docx":
        log.error(
            f"Wrong file extension: expected .docx, got {path.suffix} [pipeline:{pipeline_id}]"
        )
        raise ValueError(f"Expected a .docx file, but got: {path.suffix}")
    return path


def _build_timestamped_output_filename(save_object: OUTPUT_TYPE) -> str:
    """Apply a per-run timestamp to the output's base filename."""
    # Get pipeline ID for logs.
    pipeline_id = get_pipeline_run_id()

    # Get the base filename string
    if isinstance(save_object, document.Document):
        save_filename = constants.OUTPUT_DOCX_FILENAME
    elif isinstance(save_object, presentation.Presentation):
        save_filename = constants.OUTPUT_PPTX_FILENAME
    else:
        log.error(
            f"Unexpected output object type passed to _build_timestamped_output_filename(): {save_object}. Only document.Document or presentation.Presentation objects are supported. [pipeline:{pipeline_id}]"
        )
        raise RuntimeError(f"Unexpected output object type: {save_object}")

    # Add a timestamp to the filename
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    name, ext = save_filename.rsplit(
        ".", 1
    )  # The 1 is telling rsplit() to split from the right side and do a maximum of 1 split.
    timestamped_filename = f"{name}_{timestamp}.{ext}"

    return timestamped_filename


# endregion


# region Disk I/O - Write
def save_output(save_object: OUTPUT_TYPE, cfg: UserConfig) -> None:
    """Save the generated output object to disk as a file. Genericized to output either docx or pptx depending on which pipeline is running."""
    pipeline_id = get_pipeline_run_id()

    # Get the output folder from the config object
    save_folder = cfg.get_output_folder()

    # Apply a timestamp to the base filename
    timestamped_filename = _build_timestamped_output_filename(save_object)

    # Report if the content we're about to save is excessively huge
    _validate_content_size(save_object)

    # Create the output folder if we need to
    save_folder.mkdir(parents=True, exist_ok=True)

    output_filepath = save_folder / timestamped_filename

    # Attempt to save
    try:
        save_object.save(str(output_filepath))
        log.info(f"Successfully saved to {output_filepath}. [pipeline:{pipeline_id}]")
    except PermissionError as e:
        log.error(f"Save failed due to permission error [pipeline:{pipeline_id}]: {e}")
        raise PermissionError("Save failed: File may be open in another program")
    except OSError as e:
        log.error(f"Save failed in [pipeline:{pipeline_id}]: {e}")
        raise OSError(f"Save failed (disk space or IO issue): {e}")
    except Exception as e:
        log.error(f"Save failed in [pipeline:{pipeline_id}]: {e}")
        raise RuntimeError(f"Save failed with unexpected error: {e}")


# TODO, leafy: I'd really like `_validate_content_size()` to validate we're not about to save 100MB+ files. But that's not easy to
# estimate from the runtime object.
# For now we'll check for absolutely insane slide or paragraph counts, and just report it to the
# debug/logger.


def _validate_content_size(save_object: OUTPUT_TYPE) -> None:
    """Report if the output content we're about to save is excessively large."""
    if isinstance(save_object, document.Document):
        max_p_count = 10000
        if len(save_object.paragraphs) > max_p_count:
            log.warning(
                f"This is about to save a docx file with over {max_p_count} paragraphs ... that seems a bit long!"
            )
    elif isinstance(save_object, presentation.Presentation):
        max_s_count = 1000
        if len(list(save_object.slides)) > max_s_count:
            log.warning(
                f"This is about to save a pptx file with over {max_s_count} slides ... that seems a bit long!"
            )


# endregion


# region Disk I/O - Read & Validate
def load_and_validate_docx(input_filepath: Path) -> document.Document:
    """Use python-docx to read in the docx file contents and store to a runtime variable."""
    pipeline_id = get_pipeline_run_id()

    # Try to load the docx
    try:
        doc = docx.Document(input_filepath)  # type: ignore
    except Exception as e:
        log.error(
            f"Could not load document {str(input_filepath)} [pipeline:{pipeline_id}]. Error: {e} "
        )
        raise ValueError(f"Document appears to be corrupted: {e}")

    # Validate it contains content
    if not doc.paragraphs:
        log.error(
            f"Document {str(input_filepath)} contains no paragraphs [pipeline:{pipeline_id}]"
        )
        raise ValueError("Document contains no paragraphs.")

    first_para_w_text = _find_first_docx_paragraph_with_text(doc.paragraphs)
    if first_para_w_text is None:
        log.error(
            f"Document {str(input_filepath)} contains no text content [pipeline:{pipeline_id}]"
        )
        raise ValueError(
            "Document contains no text content, so there's nothing for the pipeline to do."
        )

    # Report content information to the user
    paragraph_count = len(doc.paragraphs)
    log.info(
        f"This document has {paragraph_count} paragraphs in it. [pipeline:{pipeline_id}]"
    )

    text = first_para_w_text.text
    preview = text[:20] + ("..." if len(text) > 20 else "")
    log.info(
        f"The first paragraph containing text begins with: {preview}. [pipeline:{pipeline_id}]"
    )

    return doc


def load_and_validate_pptx(pptx_path: Path | str) -> presentation.Presentation:
    """
    Read in pptx file contents, validate minimum content is present, and store to a runtime object. (pptx2docx-text pipeline)
    """
    pipeline_id = get_pipeline_run_id()
    # Try to load the pptx
    try:
        prs = pptx.Presentation(str(pptx_path))
    except Exception as e:
        log.error(
            f"Could not load PowerPoint file {str(pptx_path)} [pipeline:{pipeline_id}]. Error: {e} "
        )
        raise ValueError(f"Presentation appears to be corrupted: {e}")

    # Validate the pptx contains slides, and at least one contains content.
    if not prs.slides:
        log.error(
            f"Document {str(pptx_path)} contains no slides. [pipeline:{pipeline_id}]"
        )
        raise ValueError("Presentation contains no slides.")

    first_slide = _find_first_slide_with_text(list(prs.slides))
    if first_slide is None:
        log.error(
            f"No slides in {str(pptx_path)} contain text content. [pipeline:{pipeline_id}]"
        )
        raise ValueError(
            f"No slides in {str(pptx_path)} contain text content, so there's nothing for the pipeline to do."
        )

    # Report content information to the user
    slide_count = len(prs.slides)
    log.info(
        f"The pptx file {pptx_path} has {slide_count} slide(s) in it. [pipeline:{pipeline_id}]"
    )

    first_slide_paragraphs = get_slide_paragraphs(first_slide)

    log.info(
        f"The first slide detected with text content is slide_id: {first_slide.slide_id} (inside presentation.xml). [pipeline:{pipeline_id}]"
    )

    for p in first_slide_paragraphs:
        if p.text.strip():
            text = p.text.strip()
            preview = text[:20] + ("..." if len(text) > 20 else "")
            log.info(f"The first paragraph containing text begins with: {preview}")
            break
    # An else on a for-loop runs if we never hit break. This is here because I'm maybe-overly defensive in programming style.
    else:
        log.warning(f"(Could not extract preview text) [pipeline:{pipeline_id}]")

    # Return the runtime object
    return prs


# endregion


# region load & validate helpers
def _find_first_docx_paragraph_with_text(
    paragraphs: list[Paragraph_docx],
) -> Paragraph_docx | None:
    """Find the first paragraph that contains any text content in a docx."""
    for paragraph in paragraphs:
        if paragraph.text and paragraph.text.strip():
            return paragraph
    return None


def _find_first_slide_with_text(slides: list[Slide]) -> Slide | None:
    """Find the first slide that contains any paragraphs with text content."""
    for slide in slides:
        if get_slide_paragraphs(slide):
            return slide
    return None


# endregion
