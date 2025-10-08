"""TODO"""

from __future__ import annotations
from docx2pptx_text import config
from docx2pptx_text.utils import debug_print, setup_console_encoding
from docx2pptx_text import pipeline_docx2pptx
from docx2pptx_text import pipeline_pptx2docx
from docx2pptx_text.internals.config.define_config import UserConfig


def main() -> None:
    """Entry point for program flow."""
    setup_console_encoding()
    debug_print("Hello, manuscript parser!")

     # Create config with defaults
    cfg = UserConfig()
    
    # TODO: Later you'll load from YAML, merge CLI args, etc.
    # For now, just use defaults
    # Right now UserConfig() with no arguments will use all defaults, 
    # which should work for your existing sample workflow. Later when
    # you add YAML loading (Layer 3), you'll replace that with 
    # cfg = load_from_yaml_and_merge(...).

    # TODO remove config.INPUT_DOCX_FILE after we move it
    pipeline_docx2pptx.run_docx2pptx_pipeline(config.INPUT_DOCX_FILE, cfg)


    pipeline_pptx2docx.run_pptx2docx_pipeline(config.INPUT_PPTX_FILE, cfg)


# region call main
if __name__ == "__main__":
    main()
# endregion
