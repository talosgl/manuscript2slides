"""Application-wide constants and configuration values."""

# Slide layout used by docx2pptx pipeline when creating new slides from chunks.
# NOTE: We don't allow this to be overridden by users because of the dependency on
# the structure of the layout during the slide creation and content injection
# process. I actually question the wisdom of allowing template overrides at all;
# however, if they copy-paste the reference template and change formatting, etc.,
# that should be fine.
SLD_LAYOUT_CUSTOM_NAME = "manuscript2slides"

# Metadata headers/footers used in both pipelines when writing to/reading from slide speaker notes
METADATA_MARKER_HEADER: str = "START OF JSON METADATA FROM SOURCE DOCUMENT"
METADATA_MARKER_FOOTER: str = "END OF JSON METADATA FROM SOURCE DOCUMENT"
NOTES_MARKER_HEADER: str = "START OF COPIED NOTES FROM SOURCE DOCX"
NOTES_MARKER_FOOTER: str = "END OF COPIED NOTES FROM SOURCE DOCX"

# Output filename base which is combined with a unique timestamps on save to prevent clobbering
OUTPUT_PPTX_FILENAME = r"manuscript2slides_output.pptx"

OUTPUT_DOCX_FILENAME = r"pptx2docx-text_output.docx"

# Fallback for get_debug_mode() in utils
DEBUG_MODE_DEFAULT = False  # Hard-coded default
