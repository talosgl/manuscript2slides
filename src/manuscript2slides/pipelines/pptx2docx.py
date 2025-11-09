# pptx2docx.py
"""PowerPoint to Word conversion pipeline."""

import logging
import sys

from pptx import presentation
from manuscript2slides.templates import create_empty_document

from manuscript2slides import io
from manuscript2slides.internals.config.define_config import UserConfig
from manuscript2slides.processing.populate_docx import copy_slides_to_docx_body
from manuscript2slides.internals.run_context import start_pipeline_run
from manuscript2slides.internals.run_context import get_pipeline_run_id

log = logging.getLogger("manuscript2slides")


def run_pptx2docx_pipeline(cfg: UserConfig) -> None:
    """Orchestrates the pptx2docxtext pipeline."""

    # Get the pipeline_id for logging
    pipeline_id = get_pipeline_run_id()
    log.info(f"Starting pptx2docx pipeline [pipeline:{pipeline_id}]")

    pptx_path = cfg.get_input_pptx_file()

    # Safety check
    if pptx_path is None:
        raise ValueError(
            "pptx_path is None inside run_docx2pptx_pipeline(), somehow. This should never happen. "
            "Our Validation failed to catch missing input file. "
            "If you are trying to test something, use UserConfig.with_defaults() or UserConfig.for_demo() to create a test config."
        )

    user_prs: presentation.Presentation = io.load_and_validate_pptx(pptx_path)

    # Create an empty docx
    new_doc = create_empty_document(cfg)

    copy_slides_to_docx_body(user_prs, new_doc, cfg)

    log.debug(f"Attempting to save new docx file. [pipeline:{pipeline_id}]")

    io.save_output(new_doc, cfg)

    log.info(f"pptx2docx pipeline complete [pipeline:{pipeline_id}]")
    log.info(f"See log: {cfg.get_log_folder()}")
