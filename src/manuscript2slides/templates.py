# templates.py
"""Load docx and pptx templates from disk, validate shape, and create in-memory python objects from them."""
from pathlib import Path

import pptx
from pptx import presentation
import docx
from docx import document

from manuscript2slides import io
from manuscript2slides.internals import constants
from manuscript2slides.internals.config.define_config import UserConfig

# import logging
# log = logging.getLogger("manuscript2slides")


def create_empty_slide_deck(cfg: UserConfig) -> presentation.Presentation:
    """Load the PowerPoint template, create a new presentation object, and validate it contains the custom layout. (manuscript2slides pipeline)"""

    # Try to load the pptx
    try:
        template_path = cfg.get_template_pptx_path()
        validated_template = io.validate_pptx_path(Path(template_path))
        prs = pptx.Presentation(str(validated_template))
    except Exception as e:
        raise ValueError(f"Could not load template file (may be corrupted): {e}")

    # Validate it has the required slide layout for the pipeline
    layout_names = [layout.name for layout in prs.slide_layouts]
    if constants.SLD_LAYOUT_CUSTOM_NAME not in layout_names:
        raise ValueError(
            f"Template is missing the required layout: '{constants.SLD_LAYOUT_CUSTOM_NAME}'. "
            f"Available layouts: {', '.join(layout_names)}"
            f"If error persists, try renaming the Documents/manuscript2slides/templates/ folder to templates_old/ or deleting it."
        )

    return prs


def create_empty_document(cfg: UserConfig) -> document.Document:
    """
    Load Word template and create document object.

    Validates the template is a valid docx file.

    Raises:
        ValueError: If template is corrupted or invalid.
    """
    from manuscript2slides.io import validate_docx_path  # Avoid circular import

    try:
        template_path = cfg.get_template_docx_path()
        validated_template = validate_docx_path(Path(template_path))
        doc = docx.Document(str(validated_template))
    except Exception as e:
        raise ValueError(f"Could not load docx template (may be corrupted): {e}")

    # TODO, v1 required: Add validation here if needed (e.g., check for required styles)

    return doc
