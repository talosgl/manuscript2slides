"""TODO: write docstring to please ruff"""
# TODO: rename file to consts.py & consider moving to internals

from pathlib import Path

# region Actual CONSTs
# Get the directory where this script lives (NOT INTENDED FOR USER EDITING)
ROOT_DIR = Path(__file__).parent.parent.parent

# Slide layout used by docx2pptx pipeline when creating new slides from chunks. 
# All slides use the same layout.
SLD_LAYOUT_CUSTOM_NAME = "docx2pptx"

# Metadata headers/footers used in both pipelines when writing to/reading from slide speaker notes
METADATA_MARKER_HEADER: str = "START OF JSON METADATA FROM SOURCE DOCUMENT"
METADATA_MARKER_FOOTER: str = "END OF JSON METADATA FROM SOURCE DOCUMENT"
NOTES_MARKER_HEADER: str = "START OF COPIED NOTES FROM SOURCE DOCX"
NOTES_MARKER_FOOTER: str = "END OF COPIED NOTES FROM SOURCE DOCX"

# Output filename base which is combined with a unique timestamps on save to prevent clobbering
OUTPUT_PPTX_FILENAME = r"docx2pptx-text_output.pptx"

OUTPUT_DOCX_FILENAME = r"pptx2docx-text_output.docx"
# endregion

# Toggle on/off whether to print debug_prints() to the console
DEBUG_MODE = True
