# formatting.py
"""Formatting functions for both pipelines."""

# For python-pptx's private _Run and _Paragraph classes:
# pyright: reportPrivateUsage=false

# For incomplete type stubs in python-pptx:
# pyright: reportAttributeAccessIssue=false
# mypy: disable-error-code="import-untyped"

# region imports
import logging
import xml.etree.ElementTree as ET
from typing import Union, cast

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_UNDERLINE
from docx.opc.package import OpcPackage
from docx.shared import RGBColor as RGBColor_docx
from docx.styles.style import ParagraphStyle as ParagraphStyle_docx
from docx.text.font import Font as Font_docx
from docx.text.paragraph import Paragraph as Paragraph_docx
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run as Run_docx
from pptx.dml.color import RGBColor as RGBColor_pptx
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT as PP_ALIGN
from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE
from pptx.oxml.xmlchemy import OxmlElement as OxmlElement_pptx
from pptx.text.text import Font as Font_pptx
from pptx.text.text import _Paragraph as Paragraph_pptx
from pptx.text.text import _Run as Run_pptx
from pptx.util import Pt

from manuscript2slides.annotations.docx_xml import (
    extract_theme_fonts_from_xml,
    parse_xml_blob,
)
from manuscript2slides.internals.define_config import UserConfig

# endregion

log = logging.getLogger("manuscript2slides")


# region consts
# region colormap

COLOR_MAP_HEX = {
    WD_COLOR_INDEX.YELLOW: "FFFF00",
    WD_COLOR_INDEX.PINK: "FF00FF",
    WD_COLOR_INDEX.BLACK: "000000",
    WD_COLOR_INDEX.WHITE: "FFFFFF",
    WD_COLOR_INDEX.BLUE: "0000FF",
    WD_COLOR_INDEX.BRIGHT_GREEN: "00FF00",
    WD_COLOR_INDEX.DARK_BLUE: "000080",
    WD_COLOR_INDEX.DARK_RED: "800000",
    WD_COLOR_INDEX.DARK_YELLOW: "808000",
    WD_COLOR_INDEX.GRAY_25: "C0C0C0",
    WD_COLOR_INDEX.GRAY_50: "808080",
    WD_COLOR_INDEX.GREEN: "008000",
    WD_COLOR_INDEX.RED: "FF0000",
    WD_COLOR_INDEX.TEAL: "008080",
    WD_COLOR_INDEX.TURQUOISE: "00FFFF",
    WD_COLOR_INDEX.VIOLET: "800080",
}

COLOR_MAP_FROM_HEX = {v: k for k, v in COLOR_MAP_HEX.items()}
# endregion

# region alignment map
ALIGNMENT_MAP_WD2PP = {
    WD_ALIGN_PARAGRAPH.LEFT: PP_ALIGN.LEFT,
    WD_ALIGN_PARAGRAPH.CENTER: PP_ALIGN.CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT: PP_ALIGN.RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: PP_ALIGN.JUSTIFY,
    WD_ALIGN_PARAGRAPH.DISTRIBUTE: PP_ALIGN.DISTRIBUTE,
    # I don't know Thai and I'm completely guessing that this is desired
    WD_ALIGN_PARAGRAPH.THAI_JUSTIFY: PP_ALIGN.THAI_DISTRIBUTE,
    # Word has multiple JUSTIFY variants with different character spacing;
    # PowerPoint only has standard JUSTIFY and JUSTIFY_LOW
    WD_ALIGN_PARAGRAPH.JUSTIFY_HI: PP_ALIGN.JUSTIFY,
    WD_ALIGN_PARAGRAPH.JUSTIFY_MED: PP_ALIGN.JUSTIFY,
    WD_ALIGN_PARAGRAPH.JUSTIFY_LOW: PP_ALIGN.JUSTIFY_LOW,
}

# Reverse map: PP_ALIGN -> WD_ALIGN_PARAGRAPH
# Note: Where multiple WD values map to one PP value, the last entry in
# ALIGNMENT_MAP_WD2PP wins (e.g., PP_ALIGN.JUSTIFY -> WD_ALIGN_PARAGRAPH.JUSTIFY_LOW)
ALIGNMENT_MAP_PP2WD = {v: k for k, v in ALIGNMENT_MAP_WD2PP.items()}
# endregion

# region underline map
# Map WD_UNDERLINE enum values to MSO_TEXT_UNDERLINE_TYPE for docx->pptx conversion
UNDERLINE_MAP_WD2MSO = {
    WD_UNDERLINE.SINGLE: True,  # Standard single underline
    WD_UNDERLINE.DOUBLE: MSO_TEXT_UNDERLINE_TYPE.DOUBLE_LINE,
    WD_UNDERLINE.THICK: MSO_TEXT_UNDERLINE_TYPE.HEAVY_LINE,
    WD_UNDERLINE.DOTTED: MSO_TEXT_UNDERLINE_TYPE.DOTTED_LINE,
    WD_UNDERLINE.DASH: MSO_TEXT_UNDERLINE_TYPE.DASH_LINE,
    WD_UNDERLINE.DOT_DASH: MSO_TEXT_UNDERLINE_TYPE.DOT_DASH_LINE,
    WD_UNDERLINE.DOT_DOT_DASH: MSO_TEXT_UNDERLINE_TYPE.DOT_DOT_DASH_LINE,
    WD_UNDERLINE.WAVY: MSO_TEXT_UNDERLINE_TYPE.WAVY_LINE,
    WD_UNDERLINE.WAVY_DOUBLE: MSO_TEXT_UNDERLINE_TYPE.WAVY_DOUBLE_LINE,
    WD_UNDERLINE.WORDS: True,  # Word-only underline -> single underline
}
# endregion

BASELINE_SUBSCRIPT_SMALL_FONT = -25000
BASELINE_SUBSCRIPT_LARGE_FONT = -50000
BASELINE_SUPERSCRIPT_SMALL_FONT = 60000  # For fonts < 24pt
BASELINE_SUPERSCRIPT_LARGE_FONT = 30000  # For fonts >= 24pt
# endregion


# region shared formatting funcs


# region _copy_basic_font_formatting
def _copy_basic_font_formatting(
    source_font: Union[Font_docx, Font_pptx], target_font: Union[Font_docx, Font_pptx]
) -> None:
    """Extract common formatting logic for Runs (or Paragraphs)."""

    # If this font lives on a Paragraph_docx instead of a run, this might duplicate a step that
    # just happened in _copy_paragraph_font_name_docx2pptx, but it's okay, that duplication
    # on the paragraph-level is worth it to have this function be polymorphic
    if source_font.name is not None:
        target_font.name = source_font.name

    # Bold/Italics: Only overwrite when explicitly set on the source (avoid clobbering inheritance)
    if source_font.bold is not None:
        target_font.bold = source_font.bold
    if source_font.italic is not None:
        target_font.italic = source_font.italic

    # Underline: Handle both boolean and enum values
    if source_font.underline is not None:
        # Check if it's a boolean (True/False/None)
        if isinstance(source_font.underline, bool):
            target_font.underline = source_font.underline
        else:
            # It's a WD_UNDERLINE enum - map to MSO_TEXT_UNDERLINE_TYPE
            # Use mapped value if available, otherwise fall back to simple boolean
            target_font.underline = UNDERLINE_MAP_WD2MSO.get(
                source_font.underline, bool(source_font.underline)
            )


# endregion


# region _copy_font_size_formatting
def _copy_font_size_formatting(
    source_font: Union[Font_docx, Font_pptx], target_font: Union[Font_docx, Font_pptx]
) -> None:
    if source_font.size is not None:
        target_font.size = Pt(source_font.size.pt)
        """
        <a:r>
            <a:rPr lang="en-US" sz="8800" i="1" dirty="0"/>
            <a:t>MAKE this text BIG!</a:t>
        </a:r>
        """


# endregion


# region _copy_font_color_formatting
def _copy_font_color_formatting(
    source_font: Union[Font_docx, Font_pptx], target_font: Union[Font_docx, Font_pptx]
) -> None:
    # Color: copy only if source has an explicit RGB
    src_rgb = getattr(getattr(source_font, "color", None), "rgb", None)
    if src_rgb is not None:
        if isinstance(target_font, Font_pptx):
            target_font.color.rgb = RGBColor_pptx(*src_rgb)
        elif isinstance(target_font, Font_docx):
            target_font.color.rgb = RGBColor_docx(*src_rgb)


# endregion


# region _exp_fmt_issue helper
def _exp_fmt_issue(formatting_type: str, run_text: str, e: Exception) -> str:
    """Construct error message string per experimental formatting type."""
    message = f"We found a {formatting_type} in the experimental formatting JSON from a previous docx2pptx run, but we couldn't apply it. \n Run text: {run_text[:50]}... \n Error: {e}"
    return message


# endregion

# endregion


# region get docx2pptx formatting


# region copy_run_formatting_docx2pptx
def copy_run_formatting_docx2pptx(
    source_run: Run_docx,
    target_run: Run_pptx,
    experimental_formatting_metadata: list,
    cfg: UserConfig,
) -> None:
    """Mutates a pptx _Run object to apply text and formatting from a docx Run object."""
    sfont = source_run.font
    tfont = target_run.font

    target_run.text = source_run.text

    _copy_basic_font_formatting(sfont, tfont)

    _copy_font_size_formatting(sfont, tfont)

    _copy_font_color_formatting(sfont, tfont)

    if cfg.experimental_formatting_on:
        if source_run.text and source_run.text.strip():
            _copy_experimental_formatting_docx2pptx(
                source_run, target_run, experimental_formatting_metadata
            )


# endregion


# region _copy_experimental_formatting_docx2pptx
def _copy_experimental_formatting_docx2pptx(
    source_run: Run_docx,
    target_run: Run_pptx,
    experimental_formatting_metadata: list,
) -> None:
    """
    Extract experimental formatting from the docx Run and attempt to apply it to the pptx run. Additionally,
    store the formatting information in a metadata list (for the purpose of saving to JSON and enabling restoration
    during the reverse pipeline).
    """

    sfont = source_run.font
    tfont = target_run.font

    # The following code, which extends formatting support beyond python-pptx's capabilities,
    # is adapted from the md2pptx project, particularly from ./paragraph.py
    # Original source: https://github.com/MartinPacker/md2pptx
    # Author: Martin Packer
    # License: MIT
    try:
        if sfont.highlight_color is not None:
            experimental_formatting_metadata.append(
                {
                    "ref_text": source_run.text,
                    "highlight_color_enum": sfont.highlight_color.name,
                    "formatting_type": "highlight",
                }
            )
            try:
                # Convert the docx run highlight color to a hex string
                tfont_hex_str = COLOR_MAP_HEX.get(sfont.highlight_color)

                # Create an object to represent this run in memory
                rPr = target_run._r.get_or_add_rPr()

                # Create a highlight Oxml object in memory
                hl = OxmlElement_pptx("a:highlight")

                # Create a srgbClr Oxml object in memory
                srgbClr = OxmlElement_pptx("a:srgbClr")

                # Set the attribute val of the srgbClr Oxml object in memory to the desired color
                setattr(srgbClr, "val", tfont_hex_str)

                # Add srgbClr object inside the hl Oxml object
                hl.append(srgbClr)

                # Add the hl object to the run representation object, which will add all our Oxml elements inside it
                rPr.append(hl)

            except Exception as e:

                log.warning(
                    f"We found a highlight in a docx run but couldn't apply it. \n Run text: {source_run.text[:50]}... \n Error: {e}"
                )
            """
            Reference pptx XML for highlighting:
            <a:r>
                <a:rPr>
                    <a:highlight>
                        <a:srgbClr val="FFFF00"/>
                    </a:highlight>
                </a:rPr>
                <a:t>Highlight this text.</a:t>
            </a:r>
            """

        if sfont.strike:
            experimental_formatting_metadata.append(
                {"ref_text": source_run.text, "formatting_type": "strike"}
            )
            try:
                tfont._element.set("strike", "sngStrike")
            except Exception as e:
                log.warning(
                    f"Failed to apply single-strikethrough. \nRun text: {source_run.text[:50]}... \n Error: {e}"
                )

            """
            Reference pptx XML for single strikethrough:
            <a:p>
                <a:r>
                    <a:rPr lang="en-US" strike="sngStrike" dirty="0"/>
                    <a:t>Strike this text.</a:t>
                </a:r>
            </a:p>        
            """

        if sfont.double_strike:
            experimental_formatting_metadata.append(
                {"ref_text": source_run.text, "formatting_type": "double_strike"}
            )
            try:
                tfont._element.set("strike", "dblStrike")
            except Exception as e:
                log.warning(
                    f"""
                            Failed to apply double-strikthrough.
                            \nRun text: {source_run.text[:50]}... \n Error: {e}
                            \nWe'll attempt single strikethrough."""
                )
                tfont._element.set("strike", "sngStrike")
            """
            Reference pptx XML for double strikethrough:
            <a:p>
                <a:r>
                    <a:rPr lang="en-US" strike="dblStrike" dirty="0" err="1"/>
                    <a:t>Double strike this text.</a:t>
                </a:r>        
            </a:p>
            """

        if sfont.subscript:
            experimental_formatting_metadata.append(
                {"ref_text": source_run.text, "formatting_type": "subscript"}
            )
            try:
                if tfont.size is None or tfont.size < Pt(24):
                    # Cast to string on set; if we store the const as a string, the negative sign gets lost for some reason.
                    tfont._element.set("baseline", str(BASELINE_SUBSCRIPT_SMALL_FONT))
                else:
                    tfont._element.set("baseline", str(BASELINE_SUBSCRIPT_LARGE_FONT))

            except Exception as e:
                log.warning(
                    f"""
                            Failed to apply subscript. 
                            \nRun text: {source_run.text[:50]}... 
                            \n Error: {e}"""
                )
            """
            Reference pptx XML for subscript:
            <a:r>
                <a:rPr lang="en-US" baseline="-25000" dirty="0" err="1"/>
                <a:t>Subscripted text</a:t>
            </a:r>
            """

        if sfont.superscript:
            experimental_formatting_metadata.append(
                {"ref_text": source_run.text, "formatting_type": "superscript"}
            )
            try:
                if tfont.size is None or tfont.size < Pt(24):
                    tfont._element.set("baseline", str(BASELINE_SUPERSCRIPT_SMALL_FONT))
                else:
                    tfont._element.set("baseline", str(BASELINE_SUPERSCRIPT_LARGE_FONT))

            except Exception as e:
                log.warning(
                    f"""
                            Failed to apply superscript. 
                            \nRun text: {source_run.text[:50]}... 
                            \n Error: {e}"""
                )
            """
            Reference pptx XML for superscript
            <a:r>
                <a:rPr lang="en-US" baseline="30000" dirty="0" err="1"/>
                <a:t>Superscript this text.</a:t>
            </a:r>
            """

        # The below caps-handling code is not directly from md2pptx,
        # but is heavily influenced by it.
        if sfont.all_caps:
            experimental_formatting_metadata.append(
                {"ref_text": source_run.text, "formatting_type": "all_caps"}
            )
            try:
                tfont._element.set("cap", "all")
            except Exception as e:
                log.warning(
                    f"""
                            Failed to apply all caps. 
                            \nRun text: {source_run.text[:50]}... 
                            \n Error: {e}"""
                )
            """
            Reference XML for all caps:
            <a:r>
                <a:rPr lang="en-US" cap="all" dirty="0" err="1"/>
                <a:t>Put this text in all caps.</a:t>
            </a:r>
            """

        if sfont.small_caps:
            experimental_formatting_metadata.append(
                {"ref_text": source_run.text, "formatting_type": "small_caps"}
            )
            try:
                tfont._element.set("cap", "small")
            except Exception as e:
                log.warning(
                    f"""
                            Failed to apply small caps on run with text body: 
                            \nRun text: {source_run.text[:50]}... 
                            \n Error: {e}"""
                )
            """
            Reference pptx XML for small caps:
            <a:r>
                <a:rPr lang="en-US" cap="small" dirty="0" err="1"/>
                <a:t>Put this text in small caps.</a:t>
            </a:r>
            """
    except Exception as e:
        log.warning(f"Unexpected error in experimental formatting: {e}")


# endregion


# region get_theme_fonts_from_docx_package
def get_theme_fonts_from_docx_package(
    package: OpcPackage | None,
) -> dict[str, str | None]:
    """
    Extracts theme fonts from a document package (accessible via paragraph.part.package).
    This allows extracting theme fonts without needing the full Document object.
    """
    if package is None:
        return {"Major": None, "Minor": None}

    try:
        # Find theme parts in the package
        theme_parts = []
        for part in package.parts:
            if "theme" in str(part.partname):
                theme_parts.append(part)

        if not theme_parts:
            log.debug("No theme parts found in package.")
            return {"Major": None, "Minor": None}

        # Get the first theme part
        theme_part = theme_parts[0]
        theme_xml_blob = theme_part.blob

        # Parse and extract fonts
        theme_root = parse_xml_blob(theme_xml_blob)
        return extract_theme_fonts_from_xml(theme_root)

    except Exception as e:
        log.debug(f"Could not extract theme fonts from package: {e}")
        return {"Major": None, "Minor": None}


# endregion


# region get_style_font_name_with_fallback_docx
def get_effective_font_name_docx(
    style: ParagraphStyle_docx, theme_fonts: dict[str, str | None] | None = None
) -> str | None:
    """
    Gets the font name from a paragraph style, traversing the style hierarchy.

    Priority order:
    1. Theme fonts on the current style (highest priority - overrides explicit fonts in base styles)
    2. Explicit font name on current style
    3. Theme fonts on base styles
    4. Explicit font names on base styles (E.g., Heading1 inheriting from Normal, and using Normal's font)

    Args:
        style: The paragraph style to get the font name from
        theme_fonts: Optional dict with 'Major' and 'Minor' theme font names

    Returns the resolved font name, or None if no font is found.
    """
    current_style = style
    while current_style is not None:
        # Check theme fonts FIRST - they have priority over explicit fonts in base styles
        if theme_fonts:
            theme_font = _resolve_theme_font_docx(current_style, theme_fonts)
            if theme_font:
                return theme_font

        # Then check for explicit font names
        if current_style.font.name:
            # Found an explicit font name in this style or a base style
            return current_style.font.name

        # Move up to the base style (cast helps Pylance here if it complains)
        current_style = cast(ParagraphStyle_docx, current_style.base_style)

    # If the entire chain returns None, return None
    return None


def _resolve_theme_font_docx(
    style: ParagraphStyle_docx, theme_fonts: dict[str, str | None]
) -> str | None:
    """
    Attempt to resolve a theme font reference to an actual font name.

    Checks the style's rPr/rFonts element for theme font attributes (like asciiTheme, hAnsiTheme)
    and maps them to the actual font names from theme_fonts dict.
    """
    try:
        # Access the style's XML element
        style_element = style._element

        # Find the rFonts element within rPr (run properties)
        # Namespace for Word markup
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        rfonts = style_element.find(".//w:rPr/w:rFonts", ns)

        if rfonts is not None:
            # Check for various theme font attributes
            # The 'asciiTheme' and 'hAnsiTheme' attributes indicate which theme font to use
            ascii_theme = rfonts.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}asciiTheme"
            )
            hansi_theme = rfonts.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsiTheme"
            )

            # Common theme font references in Word:
            # - "majorHAnsi" or "majorAscii" -> Heading font (Major)
            # - "minorHAnsi" or "minorAscii" -> Body font (Minor)
            theme_ref = ascii_theme or hansi_theme

            if theme_ref:
                if "major" in theme_ref.lower():
                    return theme_fonts.get("Major")
                elif "minor" in theme_ref.lower():
                    return theme_fonts.get("Minor")

    except Exception as e:
        log.debug(f"Could not resolve theme font for style '{style.name}': {e}")

    return None


# endregion


# region copy_paragraph_formatting_docx2pptx
def copy_paragraph_formatting_docx2pptx(
    source_para: Paragraph_docx,
    target_para: Paragraph_pptx,
) -> None:
    """Copy docx paragraph font name (if set explicitly), alignment, and basics like bold, italics, etc. to a pptx paragraph."""

    # Font name - now works for both explicit fonts and theme fonts
    _copy_paragraph_font_name_docx2pptx(source_para, target_para)

    _copy_paragraph_alignment_docx2pptx(source_para, target_para)

    if source_para.style:
        # _copy_paragraph_format_docx2pptx(source_para, target_para)
        _copy_basic_font_formatting(source_para.style.font, target_para.font)

        # Copy size only for Heading styles to preserve semantic sizing without breaking auto-size
        # For body text, PowerPoint's auto-sizing works better without explicit size constraints
        is_heading = source_para.style.name and source_para.style.name.startswith(
            "Heading"
        )
        if is_heading:
            _copy_font_size_formatting(source_para.style.font, target_para.font)

        _copy_font_color_formatting(source_para.style.font, target_para.font)


# endregion


# region _copy_paragraph_font_name_docx2pptx
def _copy_paragraph_font_name_docx2pptx(
    source_para: Paragraph_docx,
    target_para: Paragraph_pptx,
) -> None:
    """Copy paragraph font name, resolving theme fonts if necessary."""
    name = None
    if source_para.style:
        # Extract theme fonts from the package for resolution
        theme_fonts = get_theme_fonts_from_docx_package(source_para.part.package)
        name = get_effective_font_name_docx(source_para.style, theme_fonts)

    if name:
        target_para.font.name = name


# endregion


# region _copy_paragraph_alignment_docx2pptx
def _copy_paragraph_alignment_docx2pptx(
    source_para: Paragraph_docx, target_para: Paragraph_pptx
) -> None:

    # 1. Start by setting the alignment based on the STYLE's definition (Lower Priority/Default)
    if source_para.style and source_para.style.paragraph_format.alignment:  # type: ignore
        target_para.alignment = ALIGNMENT_MAP_WD2PP.get(
            source_para.style.paragraph_format.alignment
        )

    # 2. OVERWRITE that value IF direct formatting was applied (Highest Priority)
    if source_para.alignment:
        # Use the map to get the correct PPTX enum for the DOCX value
        target_para.alignment = ALIGNMENT_MAP_WD2PP.get(source_para.alignment)


# endregion

# endregion


# region get pptx2docx formatting


# region copy_paragraph_formatting_pptx2docx
def copy_paragraph_formatting_pptx2docx(
    source_para: Paragraph_pptx, target_para: Paragraph_docx
) -> None:
    """Copy pptx paragraph font name, alignment, and basics like bold, italics, etc. to a pptx paragraph."""
    if (
        source_para.alignment
        and ALIGNMENT_MAP_PP2WD.get(source_para.alignment) is not None
    ):
        alignment_value = ALIGNMENT_MAP_PP2WD.get(source_para.alignment)
        if alignment_value is not None:
            target_para.alignment = alignment_value

    if source_para.font.name:
        target_para.style.font.name = source_para.font.name  # type: ignore
    else:
        slide_layout_first_font = get_effective_font_name_pptx(source_para)
        if slide_layout_first_font:
            target_para.style.font.name = slide_layout_first_font  # type: ignore


# endregion


# region copy_run_formatting_pptx2docx
def copy_run_formatting_pptx2docx(
    source_run: Run_pptx, target_run: Run_docx, cfg: UserConfig
) -> None:
    """Mutates a docx Run object to apply text and formatting from a pptx _Run object."""
    sfont = source_run.font
    tfont = target_run.font

    target_run.text = source_run.text

    _copy_basic_font_formatting(sfont, tfont)

    _copy_font_size_formatting(sfont, tfont)

    _copy_font_color_formatting(sfont, tfont)

    if source_run.text and source_run.text.strip() and cfg.experimental_formatting_on:
        _copy_experimental_formatting_pptx2docx(source_run, target_run)


# endregion


# region _copy_experimental_formatting_pptx2docx
def _copy_experimental_formatting_pptx2docx(
    source_run: Run_pptx, target_run: Run_docx
) -> None:
    """
    Extract experimental formatting from the pptx _Run and attempt to apply it to the docx Run.
    (Unlike in the docx2pptx pipeline, we don't additionally store this as metadata anywhere.)
    """
    sfont = source_run.font
    tfont = target_run.font

    try:
        sfont_xml = sfont._element.xml

        # Quick string checks before parsing
        if (
            "strike=" not in sfont_xml
            and "baseline=" not in sfont_xml
            and "cap=" not in sfont_xml
            and "a:highlight" not in sfont_xml
        ):
            return  # No experimental formatting to apply

        root = ET.fromstring(sfont_xml)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

        # Check for highlight nested element
        highlight = root.find(".//a:highlight/a:srgbClr", ns)
        if highlight is not None:
            log.debug(f"Found highlight in pptx run: {source_run.text[:30]}...")
            # Extract the color HEX out of the XML
            hex_color = highlight.get("val")
            if hex_color:
                # Convert the hex using the map map
                color_index = COLOR_MAP_FROM_HEX.get(hex_color)
                if color_index:
                    target_run.font.highlight_color = color_index

        # Check for strike/double-strike attribute
        strike = root.get("strike")
        if strike:
            if strike == "sngStrike":
                tfont.strike = True
            elif strike == "dblStrike":
                tfont.double_strike = True

        # Check for super/subscript
        baseline = root.get("baseline")
        if baseline:
            baseline_val = int(baseline)
            if baseline_val < 0:
                tfont.subscript = True
            elif baseline_val > 0:
                tfont.superscript = True

        # Check for all/small caps
        cap = root.get("cap")
        if cap:
            if cap == "all":
                tfont.all_caps = True
            elif cap == "small":
                tfont.small_caps = True

    except Exception as e:
        log.warning(
            f"Failed to parse pptx _Run formatting from XML: {e}, _Run text begins with: {source_run.text[:30]}"
        )


# endregion


# TODO: Bug? Shouldn't a user's template override any of the font settings per paragraph or run, unless a run explicitly has a typeface set separate from Normal?
# I don't think I'm handling this correctly right now. Yes, I like when my font is preserved ... but also I think it would be better to let the user control that with the template,
# not rely on the input file for the font. I think there's some conflicting stuff happening throughout with typeface. I probably need to remove all typeface


# region get_effective_font_name_pptx
def get_effective_font_name_pptx(paragraph: Paragraph_pptx) -> str | None:
    """
    Get the effective font name for a pptx paragraph.

    First checks if the paragraph has an explicit font name set. If not, searches the
    slide layouts for font definitions, prioritizing a:ea and a:cs elements (which
    typically contain actual font names) over a:latin elements (which often contain
    theme references like "+mj-lt"). Theme font references (starting with "+") are
    skipped in favor of actual font names.

    Returns the first actual font name found, or None if no font is found.

    Example XML structure:
    <a:defRPr sz="1800">
        <a:latin typeface="+mj-lt"/>
        <a:ea typeface="Arial" panose="02020602040305020204" pitchFamily="18" charset="0"/>
        <a:cs typeface="Times New Roman" panose="02020602040305020204" pitchFamily="18" charset="0"/>
    </a:defRPr>
    """
    if paragraph.font.name is not None:
        return paragraph.font.name

    try:
        slide_layouts = (
            paragraph.part.package.presentation_part.presentation.slide_layouts
        )
        for slide_layout in slide_layouts:
            # Check a:ea, a:cs, and a:latin elements for typeface attributes
            # Prioritize a:ea and a:cs which typically have actual font names
            for xpath_query in [
                ".//a:ea[@typeface]",
                ".//a:cs[@typeface]",
                ".//a:latin[@typeface]",
            ]:
                matching_elements = slide_layout.element.xpath(xpath_query)
                for el in matching_elements:
                    typeface = el.attrib.get("typeface")
                    # Skip theme font references (start with +) and return first actual font
                    if typeface and not typeface.startswith("+"):
                        return typeface
    except Exception as e:
        log.debug(
            f"Something went wrong when we tried to traverse XML to find a font name for a pptx paragraph (against our better judgment). \nException: {e}"
        )

    return None


# endregion

# endregion


# region apply_experimental_formatting_from_metadata
def apply_experimental_formatting_from_metadata(
    target_run: Run_docx, format_info: dict
) -> None:
    """Using JSON metadata from an earlier manuscript2slides run, try to restore experimental formatting metadata to a run during the reverse pipeline."""

    tfont = target_run.font
    formatting_type = format_info.get("formatting_type")

    if formatting_type == "highlight":
        highlight_enum = format_info.get("highlight_color_enum")
        if highlight_enum:
            try:
                color_index = getattr(WD_COLOR_INDEX, highlight_enum, None)
                if color_index is None:
                    log.debug(
                        f"Could not restore highlight color. Invalid enum '{highlight_enum}' in metadata for run: {target_run.text[:50]}..."
                    )
                else:
                    tfont.highlight_color = color_index
            except Exception as e:
                log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))

    elif formatting_type == "strike":
        try:
            tfont.strike = True
        except Exception as e:
            log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))

    elif formatting_type == "double_strike":
        try:
            tfont.double_strike = True
        except Exception as e:
            log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))

    elif formatting_type == "subscript":
        try:
            tfont.subscript = True
        except Exception as e:
            log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))

    elif formatting_type == "superscript":
        try:
            tfont.superscript = True
        except Exception as e:
            log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))

    elif formatting_type == "all_caps":
        try:
            tfont.all_caps = True
        except Exception as e:
            log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))

    elif formatting_type == "small_caps":
        try:
            tfont.small_caps = True
        except Exception as e:
            log.warning(_exp_fmt_issue(formatting_type, target_run.text, e))


# endregion
