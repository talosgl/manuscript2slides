# docx2pptx.py
"""Word to PowerPoint conversion pipeline."""

import logging
import sys
from manuscript2slides import io
from manuscript2slides.annotations.extract import process_chunk_annotations
from manuscript2slides.processing.chunking import create_docx_chunks
from manuscript2slides.processing.create_slides import slides_from_chunks
from manuscript2slides.internals.config.define_config import UserConfig
from manuscript2slides.templates import create_empty_slide_deck

log = logging.getLogger("manuscript2slides")


def run_docx2pptx_pipeline(cfg: UserConfig) -> None:
    """Orchestrates the docx2pptx pipeline."""

    # Validate we have what we need to run
    cfg.validate_docx2pptx_pipeline_requirements()

    user_docx = cfg.get_input_docx_file()

    # Validate it's a real path of the correct type. If it's not, return the error.
    try:
        user_docx_validated = io.validate_docx_path(user_docx)
    except FileNotFoundError:
        log.error(f"File not found: {user_docx}")
        sys.exit(1)
    except ValueError as e:
        log.error(f"{e}")
        sys.exit(1)
    except PermissionError:
        log.error(f"I don't have permission to read that file ({user_docx})!")
        sys.exit(1)

    # Load the docx file at that path.
    user_docx = io.load_and_validate_docx(user_docx_validated)

    # Chunk the docx by ___
    chunks = create_docx_chunks(user_docx, cfg.chunk_type)

    if (
        cfg.display_comments or cfg.display_footnotes or cfg.display_endnotes
    ) or cfg.preserve_docx_metadata_in_speaker_notes:
        chunks = process_chunk_annotations(chunks, user_docx, cfg)

    # Create the presentation object from template
    try:
        output_prs = create_empty_slide_deck(cfg)
    except Exception as e:
        log.error(f"Could not load template file (may be corrupted): {e}")
        sys.exit(1)

    # Mutate the presentation object by adding slides
    slides_from_chunks(output_prs, chunks, cfg)

    # Save the presentation to an actual pptx on disk
    io.save_output(output_prs, cfg)
