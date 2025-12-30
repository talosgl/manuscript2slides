"""Docx XML parsing utilities for extracting data exposed by existing interop libraries."""

import logging
import re
import xml.etree.ElementTree as ET
from typing import TypeVar

from docx import document
from docx.opc.part import Part
from docx.text.run import Run as Run_docx

from manuscript2slides.models import Endnote_docx, Footnote_docx

log = logging.getLogger("manuscript2slides")


# region parse_xml_blob
def parse_xml_blob(xml_blob: bytes | str) -> ET.Element:
    """Parse an XML blob into a string, from bytes."""
    try:
        if isinstance(xml_blob, str):
            xml_string = xml_blob
        else:
            # If footnote_blob is in bytes, or is bytes-like,
            # convert it to a string
            xml_string = bytes(xml_blob).decode("utf-8")

        # Create an ElementTree object by deserializing the footnotes.xml contents into a Python object
        root: ET.Element = ET.fromstring(xml_string)

        return root
    except UnicodeDecodeError as e:
        log.error(f"Invalid encoding in XML blob: {e}")
        raise ValueError(f"XML has invalid encoding: {e}") from e
    except ET.ParseError as e:
        log.error(f"Malformed XML: {e}")
        raise ValueError(f"XML is malformed: {e}") from e


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


# region find_theme_fonts
def extract_theme_fonts_from_xml(root: ET.Element) -> dict[str, str | None]:
    """Extracts major and minor font typeface names from the theme XML root."""

    # Define the namespace for DrawingML elements where fonts live
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    fonts: dict[str, str | None] = {"Major": None, "Minor": None}

    # Find the fontScheme element
    font_scheme = root.find(".//a:fontScheme", ns)

    if font_scheme is not None:
        # Find Major (Headings) font
        major_font = font_scheme.find("a:majorFont/a:latin", ns)
        if major_font is not None:
            fonts["Major"] = major_font.get("typeface")
            log.debug(f"Found major theme font: {fonts['Major']}")

        # Find Minor (Body/Normal) font
        minor_font = font_scheme.find("a:minorFont/a:latin", ns)
        if minor_font is not None:
            fonts["Minor"] = minor_font.get("typeface")
            log.debug(f"Found minor theme font: {fonts['Minor']}")

    return fonts
