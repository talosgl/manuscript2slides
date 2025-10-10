"""Main entry point for program flow."""

from __future__ import annotations

from manuscript2slides.utils import setup_console_encoding
from manuscript2slides import pipeline_docx2pptx
from manuscript2slides import pipeline_pptx2docx
from manuscript2slides.internals.config.define_config import UserConfig
from manuscript2slides.internals.logger import setup_logger
from manuscript2slides.internals.constants import DEBUG_MODE
from manuscript2slides.internals.scaffold import ensure_user_scaffold


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
    # TODO: I think, later, replace with some kind of router/orchestrator that goes to the right pipeline based on UI selections/config info
    pipeline_docx2pptx.run_docx2pptx_pipeline(cfg)

    pipeline_pptx2docx.run_pptx2docx_pipeline(cfg)
    # ===


# region call main
if __name__ == "__main__":
    main()
# endregion
