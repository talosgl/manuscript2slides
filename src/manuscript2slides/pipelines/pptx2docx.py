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

log = logging.getLogger("manuscript2slides")


def run_pptx2docx_pipeline(cfg: UserConfig) -> None:
    """Orchestrates the pptx2docxtext pipeline."""

    # Validate we have what we need to run this pipeline.
    cfg.validate_pptx2docx_pipeline_requirements()

    pipeline_id = start_pipeline_run()
    log.info(f"Starting pptx2docx pipeline [pipeline:{pipeline_id}]")

    pptx_path = cfg.get_input_pptx_file()

    # Safety check
    if pptx_path is None:
        raise ValueError(
            "pptx_path is None inside run_docx2pptx_pipeline(), somehow. This should never happen. "
            "Our Validation failed to catch missing input file. "
            "If you are trying to test something, use UserConfig.with_defaults() or UserConfig.for_demo() to create a test config."
        )

    # Validate the user's pptx filepath
    try:
        validated_pptx_path = io.validate_pptx_path(pptx_path)
    except FileNotFoundError:
        log.error(f"File not found: {pptx_path} [pipeline:{pipeline_id}]")
        sys.exit(1)
    except ValueError as e:
        log.error(f"{e} [pipeline:{pipeline_id}]")
        sys.exit(1)
    except PermissionError:
        log.error(
            f"I don't have permission to read that file ({pptx_path})! [pipeline:{pipeline_id}]"
        )
        sys.exit(1)

    # Load the pptx at that validated filepath
    try:
        user_prs: presentation.Presentation = io.load_and_validate_pptx(
            validated_pptx_path
        )
    except Exception as e:
        log.error(
            f"Content of powerpoint file invalid for pptx2docxtext pipeline run. Error: {e}. [pipeline:{pipeline_id}]"
        )
        sys.exit(1)

    # Create an empty docx
    new_doc = create_empty_document(cfg)

    copy_slides_to_docx_body(user_prs, new_doc, cfg)

    log.debug(f"Attempting to save new docx file. [pipeline:{pipeline_id}]")

    io.save_output(new_doc, cfg)

    log.info(f"pptx2docx pipeline complete [pipeline:{pipeline_id}]")
