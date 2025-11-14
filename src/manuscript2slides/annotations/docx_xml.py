"""XML parsing utilities for extracting annotations from docx files."""

import logging
import re
import xml.etree.ElementTree as ET
from typing import TypeVar

from docx import document
from docx.opc.part import Part
from docx.text.run import Run as Run_docx

from manuscript2slides.models import Endnote_docx, Footnote_docx

log = logging.getLogger("manuscript2slides")
NOTE_TYPE = TypeVar("NOTE_TYPE", Footnote_docx, Endnote_docx)


# region parse_xml_blob
def parse_xml_blob(xml_blob: bytes | str) -> ET.Element:
    """Parse an XML blob into a string, from bytes."""
    if isinstance(xml_blob, str):
        xml_string = xml_blob
    else:
        # If footnote_blob is in bytes, or is bytes-like,
        # convert it to a string
        xml_string = bytes(xml_blob).decode("utf-8")

    # Create an ElementTree object by deserializing the footnotes.xml contents into a Python object
    root: ET.Element = ET.fromstring(xml_string)

    return root


# endregion


# region find_xml_parts
def find_xml_parts(doc: document.Document, part_name: str) -> list[Part]:
    """Find XML parts matching the given name (e.g., 'footnotes.xml')"""
    # The zip package inspection logic
    # Inspect the docx package as a zip
    zip_package = doc.part.package

    if zip_package is None:
        log.warning("Could not access docx package.")
        return []

    part_name_parts: list[Part] = []
    for part in zip_package.parts:
        if part_name in str(part.partname):
            log.debug(f"We found a {part_name} part!")
            part_name_parts.append(part)

    return part_name_parts


# endregion


# region extract_notes_from_xml
def extract_notes_from_xml(
    root: ET.Element, note_class: type[NOTE_TYPE]
) -> dict[str, NOTE_TYPE]:
    """Extract footnotes or endnotes from XML, depending on note_class provided."""

    # Construct the strings we need to use in the XML search.
    # First, define the prefix and the namespace to which it will refer.
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # Second, construct the uri as a lookup in that dict to match how the XML works
    namespace_uri = ns["w"]

    # Third, construct the actual lookup strings. These are the full attribute name we're looking for in the data structure.
    # We must use double-curly braces to indicate we want a real curly brace in the node string.
    # And we also need an outer curly brace pair for the f-string syntax. That's why there's 3 total.
    id_attribute = f"{{{namespace_uri}}}id"  # "{http://...}id"
    type_attribute = f"{{{namespace_uri}}}type"

    notes_dict: dict[str, NOTE_TYPE] = {}

    for note in root:
        note_id = note.get(id_attribute)
        note_type = note.get(type_attribute)

        if note_id is None or note_type in ["separator", "continuationSeparator"]:
            continue

        note_full_text = "".join(note.itertext())

        # Strip leading footnote/endnote number (e.g. "1. text" -> "text") or empty period
        # ". text" if number is missing
        note_full_text = re.sub(r"^\d*\.\s*", "", note_full_text)

        note_hyperlinks = extract_hyperlinks_from_note(note)

        note_obj = note_class(
            note_id, text_body=note_full_text, hyperlinks=note_hyperlinks
        )

        notes_dict[note_id] = note_obj

    return notes_dict


# endregion


# region extract_hyperlinks_from_note
def extract_hyperlinks_from_note(element: ET.Element) -> list[str]:
    """Extract all hyperlinks from a docx footnote element."""
    hyperlinks: list[str] = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    for hyperlink in element.findall(".//w:hyperlink", ns):
        # Get the link text
        link_text = "".join([t.text or "" for t in hyperlink.findall(".//w:t", ns)])
        if link_text.strip():
            hyperlinks.append(link_text.strip())

    return hyperlinks


# endregion


# region detect_field_code_hyperlinks
def detect_field_code_hyperlinks(run: Run_docx) -> None | str:
    """
    Detect if this docx Run has a field code for instrText and it begins with HYPERLINK.
    If so, report it to the user, because we do not handle adding these to the pptx output.
    """
    try:
        run_xml: str = run.element.xml
        if "instrText" not in run_xml or "HYPERLINK" not in run_xml:
            return None
        root = ET.fromstring(run_xml)

        # Find instrText elements
        instr_texts = root.findall(
            ".//w:instrText",
            {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        )
        for instr in instr_texts:
            if instr.text and instr.text.startswith("HYPERLINK"):
                match = re.search(r'HYPERLINK\s+"([^"]+)"', instr.text)
                if match and match.group(1):
                    return match.group(1)

    except (AttributeError, ET.ParseError) as e:
        log.warning(
            f"Could not parse run XML for field codes: {e} while seeking instrText"
        )

    return None


# endregion
