"""TODO: write docstring to please ruff"""
# TODO: rename file to consts.py & consider moving to internals

from enum import Enum
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

# ========== copied to config & TODO: delete/redirect references

# The pptx file to use as the template for the slide deck
#TEMPLATE_PPTX = SCRIPT_DIR / "resources" / "blank_template.pptx"
# You can make your own template with the master slide and master notes page
# to determine how the output will look. You can customize things like font, paragraph style,
# slide size, slide layout...


# Desired output directory/folder to save the pptx in
#OUTPUT_PPTX_FOLDER = ROOT_DIR / "output"
# e.g., r"c:\my_presentations"
# If you leave it blank it'll save in the root of where you run the script from the command line

#OUTPUT_DOCX_FOLDER = ROOT_DIR / "output"
# e.g., r"c:\my_manuscripts"


# Input file to process. First, copy your docx file into the docx2slides-py/resources folder,
# then update the name at the end of the next line from "sample_doc.docx" to the real name.
# INPUT_DOCX_FILE = ROOT_DIR / "resources" / "sample_doc.docx"


# Which chunking method to use to divide the docx into slides. This enum lists the available choices:
# class ChunkType(Enum):
#     """Chunk Type Choices"""

#     HEADING_NESTED = "heading_nested"
#     HEADING_FLAT = "heading_flat"
#     PARAGRAPH = "paragraph"
#     PAGE = "page"


# And this is where to set what will be used in this run
# CHUNK_TYPE: ChunkType = ChunkType.HEADING_FLAT



# DISPLAY_COMMENTS: bool = True
DISPLAY_FOOTNOTES: bool = True
DISPLAY_ENDNOTES: bool = True


# We ought to support some way to leave speaker notes completely empty if the user really wants that, it's a valid use case.
# Documentation and tooltips should make it clear that this means metadata loss for round-trip pipeline data.
PRESERVE_DOCX_METADATA_IN_SPEAKER_NOTES: bool = True

COMMENTS_SORT_BY_DATE: bool = True
COMMENTS_KEEP_AUTHOR_AND_DATE: bool = True

EXPERIMENTAL_FORMATTING_ON: bool = True

# ========== pptx2docxtext pipeline consts

# INPUT_PPTX_FILE = (
#     ROOT_DIR / "resources" / "sample_slides_output.pptx"
# )  # "sample_slides.pptx"

#TEMPLATE_DOCX = ROOT_DIR / "resources" / "docx_template.docx"
