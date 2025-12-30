# pptx2docx.py
"""PowerPoint to Word conversion pipeline."""

# mypy: disable-error-code="import-untyped"
import logging
from pathlib import Path

from pptx import presentation

from manuscript2slides import io
from manuscript2slides.internals.define_config import UserConfig
from manuscript2slides.internals.paths import user_log_dir_path, user_output_dir
from manuscript2slides.internals.run_context import get_pipeline_run_id
from manuscript2slides.processing.populate_docx import copy_slides_to_docx_body
from manuscript2slides.templates import create_empty_document

log = logging.getLogger("manuscript2slides")


def run_pptx2docx_pipeline(cfg: UserConfig) -> Path:
    """Orchestrates the pptx2docxtext pipeline."""

    # Get the pipeline_id for logging
    pipeline_id = get_pipeline_run_id()
    log.info(f"Starting pptx2docx pipeline [pipeline:{pipeline_id}]")

    user_pptx_path = cfg.get_input_pptx_file()

    # Safety check
    if user_pptx_path is None:
        raise ValueError(
            "pptx_path is None inside run_docx2pptx_pipeline(), somehow. This should never happen. "
            "Our Validation failed to catch missing input file. "
            "If you are trying to test something, use UserConfig.with_defaults() or UserConfig.for_demo() to create a test config."
        )

    user_prs: presentation.Presentation = io.load_and_validate_pptx(user_pptx_path)

    # Create an empty docx
    new_doc = create_empty_document(cfg)

    copy_slides_to_docx_body(user_prs, new_doc, cfg)

    log.debug(f"Attempting to save new docx file. [pipeline:{pipeline_id}]")

    saved_output_path = io.save_output(new_doc, cfg)

    log.info(f"pptx2docx pipeline complete [pipeline:{pipeline_id}]")
    log.info(f"  Original: {user_pptx_path}")
    log.info(f"  -> Final:  {saved_output_path}")
    log.info(f"See log: {user_log_dir_path()}")
    log.info(f"See output folder: {(user_output_dir())}")
    return saved_output_path
