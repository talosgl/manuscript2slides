"""Main entry point for program flow."""

from __future__ import annotations

from manuscript2slides.utils import setup_console_encoding
from manuscript2slides.internals.config.define_config import (
    UserConfig,
    PipelineDirection,
    ChunkType,
)
from manuscript2slides.internals.logger import setup_logger
from manuscript2slides.internals.constants import DEBUG_MODE
from manuscript2slides.internals.scaffold import ensure_user_scaffold
from manuscript2slides.orchestrator import run_pipeline


def main() -> None:
    """Entry point for program flow."""
    setup_console_encoding()

    # Start up logging
    log = setup_logger(enable_trace=DEBUG_MODE)
    log.info("Starting manuscript2slides Log.")

    # A hello-world to probably remove someday, but I am sentimental. :)
    log.info("Hello, manuscript parser!")

    # Ensure user folders exist and templates are copied
    ensure_user_scaffold()

    # Create config with defaults
    # TODO: I think, later, add some kind of config propogation from interface: UI, YAML, CLI, etc.
    cfg = UserConfig()

    # Validate config shape
    cfg.validate()

    # === Pipeline testing

    cfg.chunk_type = ChunkType.HEADING_FLAT
    # Temporary: Run both for testing
    # cfg.direction = PipelineDirection.DOCX_TO_PPTX
    # run_pipeline(cfg)

    # cfg.direction = PipelineDirection.PPTX_TO_DOCX
    # run_pipeline(cfg)
    # TODO: Replace above with simple run_pipeline(cfg) once UI is ready

    # Temporary: Run round-trip test for development/testing
    from manuscript2slides.orchestrator import run_roundtrip_test

    run_roundtrip_test(cfg)
    # TODO: Replace above with simple run_pipeline(cfg) once UI is ready


# region call main
if __name__ == "__main__":
    main()
# endregion
