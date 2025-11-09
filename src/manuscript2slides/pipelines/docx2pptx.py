# docx2pptx.py
"""Word to PowerPoint conversion pipeline."""

import logging

from manuscript2slides import io
from manuscript2slides.annotations.extract import process_chunk_annotations
from manuscript2slides.processing.chunking import create_docx_chunks
from manuscript2slides.processing.create_slides import slides_from_chunks
from manuscript2slides.internals.config.define_config import UserConfig
from manuscript2slides.templates import create_empty_slide_deck
from manuscript2slides.internals.run_context import get_pipeline_run_id
from pathlib import Path
from manuscript2slides.internals.paths import user_log_dir_path


log = logging.getLogger("manuscript2slides")


def run_docx2pptx_pipeline(cfg: UserConfig) -> Path:
    """Orchestrates the docx2pptx pipeline."""

    # Get the pipeline_id for logging
    pipeline_id = get_pipeline_run_id()
    log.info(f"Starting docx2pptx pipeline. [pipeline:{pipeline_id}]")

    user_docx = cfg.get_input_docx_file()

    # Safety check
    if user_docx is None:
        raise ValueError(
            "user_docx is None inside run_docx2pptx_pipeline(), somehow. This should never happen. "
            "Our Validation failed to catch missing input file. "
            "If you are trying to test something, use UserConfig.with_defaults() or UserConfig.for_demo() to create a test config."
        )

    # Load the docx file at that path.
    user_docx = io.load_and_validate_docx(user_docx)

    # Chunk the docx by ___
    chunks = create_docx_chunks(user_docx, cfg.chunk_type)

    if (
        cfg.display_comments or cfg.display_footnotes or cfg.display_endnotes
    ) or cfg.preserve_docx_metadata_in_speaker_notes:
        chunks = process_chunk_annotations(chunks, user_docx, cfg)

    # Create the presentation object from template
    output_prs = create_empty_slide_deck(cfg)

    # Mutate the presentation object by adding slides
    slides_from_chunks(output_prs, chunks, cfg)

    # Save the presentation to an actual pptx on disk
    saved_output_path = io.save_output(output_prs, cfg)

    log.info(f"docx2pptx pipeline complete [pipeline:{pipeline_id}]")
    log.info(f"See log: {user_log_dir_path()}")
    return saved_output_path
