"""TODO"""
from __future__ import annotations
from src.docx2pptx_text import config
from src.docx2pptx_text.utils import debug_print, setup_console_encoding
from src.docx2pptx_text import pipeline_docx2pptx
from src.docx2pptx_text import pipeline_pptx2docx

def main() -> None:
    """Entry point for program flow."""
    setup_console_encoding()
    debug_print("Hello, manuscript parser!")

    pipeline_docx2pptx.run_docx2pptx_pipeline(config.INPUT_DOCX_FILE)

    # pipeline_pptx2docx.run_pptx2docx_pipeline(INPUT_PPTX_FILE)


# region call main
if __name__ == "__main__":
    main()
# endregion
