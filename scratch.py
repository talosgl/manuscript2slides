# pyright: basic
"""A file or me to mess about in"""

from docx2slides import open_and_load_docx, debug_print
from pathlib import Path
from docx.opc.part import Part
import sys

SCRIPT_DIR = Path(__file__).parent
INPUT_DOCX_FILE = SCRIPT_DIR / "resources" / "sample_doc.docx"

# Load the docx file at that path.
doc = open_and_load_docx(INPUT_DOCX_FILE)

# Inspect the docx package as a zip
zip_package = doc.part.package
if zip_package is None:
    debug_print("WARNING: Could not access docx package")
    sys.exit(0)  # return {}

endnote_parts: list[Part] = []
for part in zip_package.parts:
    if "endnotes" in str(part.partname):
        debug_print("We found an endnote part!")
        endnote_parts.append(part)

if not endnote_parts:
    sys.exit(0)  # return {}

endnote_blob = endnote_parts[0].blob


print(type(doc))
