"""TODO"""

from __future__ import annotations

from docx2pptx_text.utils import debug_print, setup_console_encoding
from docx2pptx_text import pipeline_docx2pptx
from docx2pptx_text import pipeline_pptx2docx
from docx2pptx_text.internals.config.define_config import UserConfig
from docx2pptx_text.internals.logger import setup_logger
from docx2pptx_text.internals.constants import DEBUG_MODE
from docx2pptx_text.internals.scaffold import ensure_user_scaffold


def main() -> None:
    """Entry point for program flow."""
    setup_console_encoding()

    debug_print("Hello, manuscript parser!")  # TODO: still need to replace all these :|

    # Start up logging
    log = setup_logger(enable_trace=DEBUG_MODE)
    log.info("Starting docx2pptx_text Log.")

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
