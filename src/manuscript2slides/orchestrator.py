"""Route program flow to the appropriate pipeline based on user-indicated direction."""

import logging
from pathlib import Path

from manuscript2slides.internals.define_config import (
    PipelineDirection,
    UserConfig,
)
from manuscript2slides.internals.manifest import RunManifest
from manuscript2slides.internals.run_context import (
    get_pipeline_run_id,
    get_session_id,
    start_pipeline_run,
)
from manuscript2slides.pipelines import docx2pptx, pptx2docx

log = logging.getLogger("manuscript2slides")


# region run_pipeline
def run_pipeline(cfg: UserConfig) -> Path:
    """Run validation and then route to the appropriate pipeline based on config."""

    cfg.pre_run_check()

    # Start pipeline run and get a fresh ID
    pipeline_id = start_pipeline_run()
    log.info(f"Initializing pipeline run. [pipeline:{pipeline_id}]")

    # Create run manifest object
    run_manifest = RunManifest(cfg, run_id=pipeline_id)
    run_manifest.start()

    # Dump pipeline run info to log
    log_pipeline_info(cfg)

    try:
        if cfg.direction == PipelineDirection.DOCX_TO_PPTX:
            output_path = docx2pptx.run_docx2pptx_pipeline(cfg)
        elif cfg.direction == PipelineDirection.PPTX_TO_DOCX:
            output_path = pptx2docx.run_pptx2docx_pipeline(cfg)
        else:
            raise ValueError(f"Unknown pipeline direction: {cfg.direction}")

        run_manifest.complete(output_path)  # Mark success
        return output_path

    except Exception as e:
        run_manifest.fail(e)  # Mark failure
        raise  # Re-raise so GUI/CLI still see the error


# endregion


# region run_round_trip_test
def run_roundtrip_test(cfg: UserConfig) -> tuple[Path, Path, Path]:
    """
    Test utility: Run both pipelines in sequence.

    Returns:
        tuple: (original_docx, intermediate_pptx, final_docx) paths for comparison
    """
    log.info("Starting round-trip test")

    # Save original input
    original_docx = cfg.get_input_docx_file()

    # Safety check
    if original_docx is None:
        log.debug(
            f"Somehow, the input_docx wasn't set before run_round_trip() was called. This shouldn't happen. "
            "It should have been called with this pre-filled by UserConfig.with_defaults() or UserConfig.for_demo() for CLI, "
            "or by user selections / config loading in the GUI."
        )
        raise ValueError(
            "Round-trip test requires input_docx to be set. "
            "Use UserConfig.with_defaults() or UserConfig.for_demo() to create a test config."
        )

    # Run docx -> pptx (direction auto-inferred from input_docx)
    run_pipeline(cfg)

    # Find the output pptx
    output_folder = cfg.get_output_folder()
    intermediate_pptx = _find_most_recent_file(output_folder, "*.pptx")
    log.info(f"Intermediate pptx: {intermediate_pptx}")

    # Create a NEW config for reverse direction
    reverse_cfg = UserConfig(
        input_pptx=intermediate_pptx,
        output_folder=cfg.output_folder,  # Keep same output location
        template_docx=cfg.template_docx,  # Keep templates
    )
    reverse_cfg.enable_all_options()  # Match the first-run bool settings

    run_pipeline(reverse_cfg)

    # Find the final output
    final_docx = _find_most_recent_file(output_folder, "*.docx")

    log.info(f"Round-trip complete:")
    log.info(f"  Original: {original_docx}")
    log.info(f"  -> PPTX:   {intermediate_pptx}")
    log.info(f"  -> Final:  {final_docx}")
    log.info(f"  Output folder:  {output_folder}")

    return original_docx, intermediate_pptx, final_docx


# endregion


# region _find_most_recent_file
def _find_most_recent_file(folder: Path, pattern: str) -> Path:
    """Find the most recently created file matching glob pattern."""
    files = list(folder.glob(pattern))
    if not files:
        log.error(f"No files matching '{pattern}' in {folder}")
        raise FileNotFoundError(
            f"Cannot proceed: No files matching '{pattern}' in {folder}"
        )

    # Use st_mtime (modification time) which is reliable across platforms
    # and makes more sense for "most recent output file"
    return max(files, key=lambda p: p.stat().st_mtime)


# endregion


# region log_pipeline_info
def log_pipeline_info(cfg: UserConfig) -> None:
    """Print this pipeline run's run ID, session ID, and general config info to the log."""
    log.info(f"=== Pipeline Run Started ===")
    log.info(f"Run ID: {get_pipeline_run_id()}")
    log.info(f"Session ID: {get_session_id()}")
    log.info(f"Direction: {cfg.direction.value}")
    log.info(f"Input: {cfg.input_docx or cfg.input_pptx}")
    log.info(f"Configuration: {cfg}")  # This will use the dataclass __repr__


# endregion
