"""Route program flow to the appropriate pipeline based on user-indicated direction."""

import logging
from pathlib import Path

from manuscript2slides.pipelines import docx2pptx, pptx2docx
from manuscript2slides.internals.config.define_config import (
    PipelineDirection,
    UserConfig,
)

log = logging.getLogger("manuscript2slides")


def run_pipeline(cfg: UserConfig) -> None:
    """Route to the appropriate pipeline based on config."""
    if cfg.direction == PipelineDirection.DOCX_TO_PPTX:
        docx2pptx.run_docx2pptx_pipeline(cfg)
    elif cfg.direction == PipelineDirection.PPTX_TO_DOCX:
        pptx2docx.run_pptx2docx_pipeline(cfg)
    else:
        raise ValueError(f"Unknown pipeline direction: {cfg.direction}")


def run_roundtrip_test(cfg: UserConfig) -> tuple[Path, Path, Path]:
    """
    Test utility: Run both pipelines in sequence.

    Returns:
        tuple: (original_docx, intermediate_pptx, final_docx) paths for comparison
    """
    log.info("Starting round-trip test")

    # Save original input
    original_docx = cfg.get_input_docx_file()

    # Run docx -> pptx
    cfg.direction = PipelineDirection.DOCX_TO_PPTX
    run_pipeline(cfg)

    # Find the output pptx
    output_folder = cfg.get_output_folder()
    intermediate_pptx = _find_most_recent_file(output_folder, "*.pptx")
    log.info(f"Intermediate pptx: {intermediate_pptx}")

    # Run pptx -> docx using the output from previous step
    cfg.input_pptx = str(intermediate_pptx)
    cfg.direction = PipelineDirection.PPTX_TO_DOCX
    run_pipeline(cfg)

    # Find the final output
    final_docx = _find_most_recent_file(output_folder, "*.docx")

    log.info(f"Round-trip complete:")
    log.info(f"  Original: {original_docx}")
    log.info(f"  -> PPTX:   {intermediate_pptx}")
    log.info(f"  -> Final:  {final_docx}")

    return original_docx, intermediate_pptx, final_docx


def _find_most_recent_file(folder: Path, pattern: str) -> Path:
    """Find the most recently created file matching glob pattern."""
    files = list(folder.glob(pattern))
    if not files:
        raise FileNotFoundError(f"No files matching '{pattern}' in {folder}")

    # Use st_mtime (modification time) which is reliable across platforms
    # and makes more sense for "most recent output file"
    return max(files, key=lambda p: p.stat().st_mtime)
