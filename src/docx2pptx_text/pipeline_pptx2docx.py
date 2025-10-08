"""TODO"""

from docx2pptx_text.utils import debug_print, setup_console_encoding
import docx
import sys
from docx2pptx_text import io
from pptx import presentation
from pathlib import Path
from docx2pptx_text.populate_docx import copy_slides_to_docx_body
from docx2pptx_text.internals.config.define_config import UserConfig


def run_pptx2docx_pipeline(cfg: UserConfig) -> None:
    """Orchestrates the pptx2docxtext pipeline."""

    pptx_path = cfg.get_input_pptx_file()
    
    # Validate the user's pptx filepath
    try:
        validated_pptx_path = io.validate_pptx_path(pptx_path)
    except FileNotFoundError:
        print(f"Error: File not found: {pptx_path}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except PermissionError:
        print(f"I don't have permission to read that file ({pptx_path})!")
        sys.exit(1)

    # Load the pptx at that validated filepath
    try:
        user_prs: presentation.Presentation = io.load_and_validate_pptx(
            validated_pptx_path
        )
    except Exception as e:
        print(
            f"Content of powerpoint file invalid for pptx2docxtext pipeline run. Error: {e}."
        )
        sys.exit(1)

    # Create an empty docx
    docx_template = cfg.get_input_docx_file()
    new_doc = docx.Document(str(docx_template))

    copy_slides_to_docx_body(user_prs, new_doc, cfg)

    debug_print("Attempting to save new docx file.")

    io.save_output(new_doc, cfg)
