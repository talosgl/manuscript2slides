"""Application-wide constants and configuration values."""

from pathlib import Path

# region Actual CONSTs
# Get the directory where this script lives (NOT INTENDED FOR USER EDITING)
PACKAGE_DIR = Path(__file__).parent.parent
RESOURCES_DIR = PACKAGE_DIR / "resources"

# Slide layout used by docx2pptx pipeline when creating new slides from chunks.
# All slides use the same layout.
# TODO, future: Allow the user to override this name? This can still be the fallback.
# There's still a lot of dependency on the structure of this layout matching what's expected
# so it may be diminishing returns to bother with a robust replacement, when they can just
# edit the existing template in pptx themselves and rename the file itself.
SLD_LAYOUT_CUSTOM_NAME = "manuscript2slides"

# Metadata headers/footers used in both pipelines when writing to/reading from slide speaker notes
METADATA_MARKER_HEADER: str = "START OF JSON METADATA FROM SOURCE DOCUMENT"
METADATA_MARKER_FOOTER: str = "END OF JSON METADATA FROM SOURCE DOCUMENT"
NOTES_MARKER_HEADER: str = "START OF COPIED NOTES FROM SOURCE DOCX"
NOTES_MARKER_FOOTER: str = "END OF COPIED NOTES FROM SOURCE DOCX"

# Output filename base which is combined with a unique timestamps on save to prevent clobbering
OUTPUT_PPTX_FILENAME = r"manuscript2slides_output.pptx"

OUTPUT_DOCX_FILENAME = r"pptx2docx-text_output.docx"
# endregion

# Toggle on/off whether to print debug_prints() to the console
# TODO, v1: allow this to be set from the UI?
DEBUG_MODE = True
DEBUG_MODE_DEFAULT = False  # Hard-coded default
SENTINEL = object()
