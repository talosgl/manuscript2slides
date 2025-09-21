# pyright: basic
"""
Convert Microsoft Word documents to PowerPoint presentations.

This tool processes .docx files and converts them into .pptx slide decks by chunking
the document content based on various strategies (paragraphs, headings, or page breaks).
Text formatting like bold, italics, and colors can optionally be preserved.

The main workflow:
1. Load a .docx file using python-docx
2. Chunk the content based on the selected strategy
3. Create slides from chunks using a PowerPoint template
4. Save the resulting .pptx file

Supported chunking methods:
- paragraph: Each paragraph becomes a slide
- page: New slide for each page break
- heading_flat: New slide for each heading (any level)
- heading_nested: New slide based on heading hierarchy

Example:
    python docx2pptx-text.py

    (Configure INPUT_DOCX_FILE and other constants before running)
"""
# region imports
from __future__ import annotations

# Standard library
import io
import platform
import re
import sys
from datetime import datetime
from enum import Enum
from pathlib import Path
from dataclasses import dataclass, field
from typing import TypeVar
import json

# import warnings

# Third-party libraries
import docx
import pptx
from docx import document
from docx.comments import Comment as Comment_docx
from docx.opc.part import Part

# from docx.table import Table as Table_docx
# from docx.text.hyperlink import Hyperlink as Hyperlink_docx
from docx.text.paragraph import Paragraph as Paragraph_docx
from docx.text.run import Run as Run_docx
from pptx import presentation
from pptx.dml.color import RGBColor as RGBColor_pptx
from pptx.slide import Slide, SlideLayout
from pptx.text.text import TextFrame, _Paragraph as Paragraph_pptx, _Run as Run_pptx  # type: ignore
from pptx.shapes.placeholder import SlidePlaceholder
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt
import xml.etree.ElementTree as ET

from pptx.slide import NotesSlide

# RUN_TYPE = TypeVar("RUN_TYPE", Run_docx, Run_pptx)
from docx.shared import RGBColor as RGBColor_docx
from docx.text.font import Font as Font_docx
from pptx.text.text import Font as Font_pptx
from typing import Union

# endregion

# region colormap
from docx.enum.text import WD_COLOR_INDEX

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
#

# region Overarching TODOs
"""
Must-Implement v0 Features:
- Reverse flow: export slide text frame content to docx paragraphs with the annotations kept as comments, inline
    - This could allow a flow where you can iterate back and forth (and remove the need to "update" an existing deck with manuscript updates)
- Change consts configuration to use a class or something
- Rearchitect to be multi-file

v1 features I'd like:
- Create Documents/docx2pptx/input/output/resources structure
    - Copies sample files from app resources to user folders
    - Cleanup mode for debug runs
    - Have the output do a rotating 5 files or so
- Build a simple UI + Package this so that any writer can use it without needing to know WTF python is
    - + Consider reworking into C# for this; good practice for me


Public v1
- Polish up enough to make repo public
- What does is "done enough for public github repo mean"? "ready when I'm comfortable having strangers use it without asking me questions." 
    - Error messages that tell users what went wrong and how to fix it
    - Code that doesn't crash on common edge cases.

Stretch Wishlist Features:
-   Add support for importing .md and .txt; split by whitespaces or character like \n\n.
-   Add support to break chunks (of any type) at a word count threshold.
-   Add support to deconstruct standard technical marketing docx demo scripts 
    (table parsing; consistent row structure (Column A = slides, Column B = notes)).

"""

"""
Known Issues & Limitations:
    -   We only support text content. No images, tables, etc., are copied between the formats, and we do not have plans 
        to support these in future.

    -   We do not support .doc or .ppt, only .docx. If you have a .doc file, convert it to .docx using Word, Google Docs, 
        or LibreOffice before processing.

    -   We do not support .ppt, only .pptx.

    -   Auto-fit text resizing in slides doesn't work. PowerPoint only applies auto-fit sizing when opened in the UI. 
        You can get around this manually with these steps:
            1. Open up the output presentation in PowerPoint Desktop > View > Slide Master
            2. Select the text frame object, right-click > Format Shape
            3. Click the Size & Properties icon {TODO ADD IMAGES}
            4. Click Text Box to see the options
            5. Toggle "Do not Autofit" and then back to "Shrink Text on overflow"
            6. Close Master View
            7. Now all the slides should have their text properly resized.

    -   Field code hyperlinks not supported - Some hyperlinks (like the sample_doc.docx's first "Where are Data?" link) 
        are stored using Word's field code format and display as plain text instead of clickable links. The exact 
        conditions that cause this format are unclear, but it may occur with hyperlinks in headings or certain copy/paste 
        scenarios. We think most normal hyperlinks will work fine. We try to report when we detect these are present but cannot
        reliably copy them as text into the body.

    -   ANNOTATIONS LIMITATIONS
        -   We collapse all comments, footnotes, and endnotes into a slide's speaker notes. PowerPoint itself doesn't 
            support real footnotes or endnotes at all. It does have a comments functionality, but the library used here 
            (python-pptx) doesn't support adding comments to slides yet. 

        -   Note that inline reference numbers (1, 2, 3, etc.) from the docx body are not preserved in the slide text - 
            only the annotation content appears in speaker notes.

        -   You can choose to preserve some comment metadata (author, timestamps) in plain text, but not threading.

        -   Footnotes and endnotes don't preserve formatting, only plain text.
    
    -   REVERSE FLOW LIMITATIONS
        -   The reverse flow (pptx2docx) is significantly less robust. 

        -   There is no support for separating out speaker notes into types of annotations. We simply take all speaker notes and attach
            them to the final paragraph of each slide's content as a comment. Original source docx comment metadata is not preserved 
            (apart from plain text author/timestamps if kept during an original docx2pptx conversion). We do not preserve any text formatting 
            for speaker notes.

        -   There will always be a blank line at the start of the reverse-pipeline document. When creating a new document with python-docx 
            using Document(), it inherently includes a single empty paragraph at the start. This behavior is mandated by the Open XML 
            .docx standard, which requires at least one paragraph within the w:body element in the document's XML structure.

        -   Powerpoint has no concept of headings. TODO: We experimentally attempt to preserve heading metadata during the docx2pptx pipeline.
            If this metadata is detected during the pptx2docx pipeline, we attempt to reapply the headings

"""
# endregion
# TODO: rename to docx2pptx-text / pptx2docx-text

# region CONSTANTS / config.py
# Get the directory where this script lives (NOT INTENDED FOR USER EDITING)
SCRIPT_DIR = Path(__file__).parent


# === docx2pptx Consts for script user to alter per-run ===

# The pptx file to use as the template for the slide deck
TEMPLATE_PPTX = SCRIPT_DIR / "resources" / "blank_template.pptx"
# You can make your own template with the master slide and master notes page
# to determine how the output will look. You can customize things like font, paragraph style,
# slide size, slide layout...

# Desired slide layout. All slides use the same layout.
SLD_LAYOUT_CUSTOM_NAME = "docx2pptx"

# Desired output directory/folder to save the pptx in
OUTPUT_PPTX_FOLDER = SCRIPT_DIR / "output"
# e.g., r"c:\my_presentations"
# If you leave it blank it'll save in the root of where you run the script from the command line

# Desired output filename; Note that this will clobber an existing file of the same name!
OUTPUT_PPTX_FILENAME = r"sample_slides_output.pptx"


# Input file to process. First, copy your docx file into the docx2slides-py/resources folder,
# then update the name at the end of the next line from "sample_doc.docx" to the real name.
INPUT_DOCX_FILE = SCRIPT_DIR / "resources" / "sample_doc.docx"


# Which chunking method to use to divide the docx into slides. This enum lists the available choices:
class ChunkType(Enum):
    """Chunk Type Choices"""

    HEADING_NESTED = "heading_nested"
    HEADING_FLAT = "heading_flat"
    PARAGRAPH = "paragraph"
    PAGE = "page"


# And this is where to set what will be used in this run
CHUNK_TYPE: ChunkType = ChunkType.HEADING_FLAT


# Toggle on/off whether to print debug_prints() to the console
DEBUG_MODE = True  # TODO, v1 POLISH: set to false before publishing; update: TODO, UX: please god replace all these consts with a config class or something for v1


DISPLAY_COMMENTS: bool = True
DISPLAY_FOOTNOTES: bool = True
DISPLAY_ENDNOTES: bool = True

DISPLAY_DOCX_ANNOTATIONS_IN_SLIDE_SPEAKER_NOTES: bool = (
    DISPLAY_COMMENTS or DISPLAY_FOOTNOTES or DISPLAY_ENDNOTES
)

# We ought to support some way to leave speaker notes completely empty if the user really wants that, it's a valid use case.
# Documentation and tooltips should make it clear that this means metadata loss for round-trip pipeline data.
PRESERVE_DOCX_METADATA_IN_SPEAKER_NOTES: bool = True

COMMENTS_SORT_BY_DATE: bool = True
COMMENTS_KEEP_AUTHOR_AND_DATE: bool = True

EXPERIMENTAL_FORMATTING_ON: bool = True

# ========== pptx2docxtext pipeline consts

INPUT_PPTX_FILE = SCRIPT_DIR / "resources" / "sample_slides.pptx"

TEMPLATE_DOCX = SCRIPT_DIR / "resources" / "docx_template.docx"

OUTPUT_DOCX_FOLDER = SCRIPT_DIR / "output"
# e.g., r"c:\my_manuscripts"

OUTPUT_DOCX_FILENAME = r"sample_pptx2docxtext_output.docx"
# endregion


# region models.py
@dataclass
class Comment_docx_custom:
    """A custom wrapper for the python-docx Comment class, allowing us to capture reference text."""

    comment_obj: Comment_docx
    reference_text: str | None = None  # The text this comment is attached to

    @property
    def note_id(self) -> int:
        """Alias for comment_id to provide a common interface with other note types."""
        return self.comment_obj.comment_id


@dataclass
class Footnote_docx:
    """
    Represents a footnote extracted from a docx.

    Contains the footnote ID, text content, and any hyperlinks found within.
    Used for preserving footnote information when python-docx doesn't provide
    direct access to footnote content.
    """

    footnote_id: str
    text_body: str
    hyperlinks: list[str] = field(default_factory=list[str])
    reference_text: str | None = None

    @property
    def note_id(self) -> str:
        """Alias for footnote_id to provide a common interface with other note types."""
        return self.footnote_id


@dataclass
class Endnote_docx:
    """
    Represents a endnote extracted from a docx.

    Contains the endnote ID, text content, and any hyperlinks found within.
    Used for preserving endnote information when python-docx doesn't provide
    direct access to endnote content.
    """

    endnote_id: str
    text_body: str
    hyperlinks: list[str] = field(default_factory=list[str])
    reference_text: str | None = None

    @property
    def note_id(self) -> str:
        """Alias for endnote_id to provide a common interface with other note types."""
        return self.endnote_id


@dataclass
class Chunk_docx:
    """Class for Chunk objects made from docx paragraphs and their associated annotations."""

    # Use "default_factory" to ensure every chunk gets its own list.
    # (Lists are mutable; it is a common error/bug to accidentally assign one list
    # shared amongst every instance of a class, rather than one per instance.)
    paragraphs: list[Paragraph_docx] = field(default_factory=list[Paragraph_docx])

    # TODO, TABLE SUPPORT:
    # If we want to support tables later, we'll need a list that has both the paragraphs
    # and tables together, in document order.
    #
    # We could use a type alias like this (outside/above the Class definition):
    # InnerContentType_docx = Paragraph_docx | Table_docx # might allow us to support tables and paragraphs
    #
    # And then inside the class definition we'd make a new list, like this:
    # inner_contents: list[InnerContentType_docx] = field(default_factory=list)
    #
    # We'd probably also need to preserve special metadata about the table for reverse-pipeline stuff.

    comments: list[Comment_docx_custom] = field(
        default_factory=list[Comment_docx_custom]
    )
    footnotes: list[Footnote_docx] = field(default_factory=list[Footnote_docx])
    endnotes: list[Endnote_docx] = field(default_factory=list[Endnote_docx])

    @classmethod
    def create_with_paragraph(cls, paragraph: Paragraph_docx) -> "Chunk_docx":
        """Create a new instance of a Chunk_docx object but also populate the paragraphs list with an initial element."""
        return cls(paragraphs=[paragraph])

    def add_paragraph(self, new_paragraph: Paragraph_docx) -> None:
        """Add a paragraph to this Chunk object's paragraphs list."""
        self.paragraphs.append(new_paragraph)

    def add_paragraphs(self, new_paragraphs: list[Paragraph_docx]) -> None:
        """Add a list of paragraphs to this Chunk object's paragraphs list."""
        self.paragraphs.extend(new_paragraphs)  # Add multiple at once

    def add_comment(self, comment: Comment_docx_custom) -> None:
        """Add a comment to this Chunk object's comment list."""
        self.comments.append(comment)

    def add_footnote(self, footnote: Footnote_docx) -> None:
        """Add a footnote to this Chunk object's footnote list."""
        self.footnotes.append(footnote)

    def add_endnote(self, endnote: Endnote_docx) -> None:
        """Add a endnote to this Chunk object's endnote list."""
        self.endnotes.append(endnote)


# TODO, REVERSE FLOW: We probably need a separate Chunk_pptx class

# I think it would be:
# list of body [Paragraph_pptx]
# list of speaker notes [Paragraph_pptx]
# make sure to use a factory for these lists because lists are mutable!

# endregion


# region __main__.py & run_pipeline.py
# eventual destination: ./src/docx2pptx/__main__.py
def main() -> None:
    """Entry point for program flow."""
    setup_console_encoding()
    debug_print("Hello, manuscript parser!")

    run_docx2pptx_pipeline(INPUT_DOCX_FILE)

    # run_pptx2docx_pipeline(INPUT_PPTX_FILE)


# endregion


# region pptx2docxtext
def run_pptx2docx_pipeline(pptx_path: Path) -> None:
    """Orchestrates the pptx2docxtext pipeline."""

    # Validate the user's pptx filepath
    try:
        validated_pptx_path = validate_pptx_path(pptx_path)
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
        user_prs: presentation.Presentation = open_and_load_pptx(validated_pptx_path)
    except Exception as e:
        print(
            f"Content of powerpoint file invalid for pptx2docxtext pipeline run. Error: {e}."
        )
        sys.exit(1)

    # Create an empty docx
    new_doc = docx.Document(str(TEMPLATE_DOCX))

    # natural language outline
    # Work sequentially through each slide in list(prs.slides), and
    # Work sequentialy through each paragraph of that slide body's text frame(s)? # TODO: Should we support multiple text frames?
    # Work sequentially through each Run_pptx of that paragraph, and...
    # Copy the run and run formatting to a Run_docx
    # (Append that run to the paragraph)
    # Next, copy everything from the slide's notes_slide text (if it exists) into a comment, and attach the comment
    # to either the beginning or end of the paragraphs we just copied into the docx

    # Pseudo... code.. ish
    # copy_slides_to_docx_body(prs, newdoc)
    # slide_list = list(prs.slides)
    # for slide in slide_list:
    # find_and_copy_all_slide_text(slide)
    # paragraphs: list() = get_slide_paragraphs(text_frames)
    # for para in paragraph:
    # TODO: is there any processing needed before processing runs, like analyzing "Run_docx" vs Hyperlink? (It doesn't seem like it)
    # copy_Run_pptx_with_formatting()
    # find_and_copy_speaker_notes(slide)

    # TODO: This is temp to make sure we can actually create shit
    # first_paragraph = new_doc.paragraphs[0] # This appears to be the only way to access the very first paragraph.
    # first_paragraph.add_run("This is the content of the first paragraph.")
    # new_para = new_doc.add_paragraph()
    # new_para.text = "Lorem Ipsum"
    # === end temp

    # Pseudo... code.. ish
    copy_slides_to_docx_body(user_prs, new_doc)
    # slide_list = list(prs.slides)
    # for slide in slide_list:
    # find_and_copy_all_slide_text(slide)
    # paragraphs: list() = get_slide_paragraphs(slide)
    # for para in paragraph:
    # TODO: is there any processing needed before processing runs, like analyzing "Run_docx" vs Hyperlink? (It doesn't seem like it)
    # TODO: apply heading styles
    # copy_Run_pptx_with_formatting()
    # find_and_copy_speaker_notes(slide)

    debug_print("Attempting to save new docx file.")
    save_output(new_doc)


def copy_slides_to_docx_body(
    prs: presentation.Presentation, new_doc: document.Document
) -> None:
    """
    For every slide in the deck, copy its main body text into the docx body, and append any slide speaker notes
    to the end of the docx paragraphs that came from that slide.
    """

    # Make a list of all slides
    slide_list = list(prs.slides)

    # For each slide...
    for slide in slide_list:

        # TODO:
        # - Restore heading metadata from speaker notes json and apply to the correct paragraph
        # - Restore comments, endnotes, footnotes from speaker notes and apply to the correct run if possible; paragraph if not.
        #       NOTE: If even the specific paragraph is undetected because of text body changes, fallback to apply it to the last
        #       paragraph from this slide, as before.

        # Get a list of this slide's paragraphs
        paragraphs: list[Paragraph_pptx] = get_slide_paragraphs(slide)

        last_run = None

        # For every paragraph, copy it to the new document
        for para in paragraphs:
            new_para = new_doc.add_paragraph()

            for run in para.runs:
                new_docx_run = new_para.add_run()
                last_run = new_docx_run
                copy_run_formatting_pptx2docx(run, new_docx_run)

        # TODO: cut this behavior/alter it to support restoreing metadata properly from the speaker notes.
        # After copying the last of this slide's runs, append the speaker notes to the last-copied run
        if (
            slide.has_notes_slide
            and slide.notes_slide.notes_text_frame is not None
            and last_run is not None
        ):
            raw_comment_text = slide.notes_slide.notes_text_frame.text
            comment_text = sanitize_xml_text(raw_comment_text)

            if comment_text.strip():
                new_doc.add_comment(last_run, comment_text)


def copy_run_formatting_pptx2docx(source_run: Run_pptx, target_run: Run_docx) -> None:
    """Mutates a docx Run object to apply text and formatting from a pptx _Run object."""
    sfont = source_run.font
    tfont = target_run.font

    target_run.text = source_run.text

    _copy_basic_run_formatting(sfont, tfont)

    _copy_run_color_formatting(sfont, tfont)

    # TODO REVERSE ALL THE EXPERIMENTAL FORMATTING STUFF


def open_and_load_pptx(pptx_path: Path | str) -> presentation.Presentation:
    """Use python-pptx to read in the pptx file contents, validate minimum content is present, and store to a runtime object."""
    prs = pptx.Presentation(str(pptx_path))

    # Count and report slides to validate we can see the content within this file.
    slide_count = len(prs.slides)
    if slide_count > 0:
        debug_print(f"The pptx file {pptx_path} has {slide_count} slide(s) in it.")

        slide_list = list(prs.slides)
        # Observations:
        # 1)    Sequential iteration order matters: slide_ids will iterate here in the order that they appear in the visual slide deck.
        # 2)    slide_id does NOT imply its place in the slide order: slide_ids appear to be n+1 generated,
        #       but a person can easily move slides around in the sequence of slides as desired. The only meaning that you can gain
        #       by sorting by slide_id is (maybe) the order in which the slide items were added to the deck, not the order that the
        #       user wants them to be viewed/read.

        first_slide = find_first_slide_with_text(slide_list)
        if first_slide is None:
            raise RuntimeError(
                f"No slides in {pptx_path} contain text content, so there's nothing for the pipeline to do."
            )

        first_slide_paragraphs = get_slide_paragraphs(first_slide)
        debug_print(
            f"The first slide detected with text content is slide_id: {first_slide.slide_id}. The text content is: \n"
        )
        for p in first_slide_paragraphs:
            debug_print(p.text)

    # Return the runtime object
    return prs


def find_first_slide_with_text(slides: list[Slide]) -> Slide | None:
    """Find the first slide that contains any paragraphs with text content."""
    for slide in slides:
        if get_slide_paragraphs(slide):
            return slide
    return None


def get_slide_paragraphs(slide: Union[Slide, NotesSlide]) -> list[Paragraph_pptx]:
    """Extract all paragraphs from all text placeholders in a slide."""
    paragraphs: list[Paragraph_pptx] = []

    for placeholder in slide.placeholders:
        if (
            isinstance(placeholder, SlidePlaceholder)
            and hasattr(placeholder, "text_frame")
            and placeholder.text_frame
        ):
            textf: TextFrame = placeholder.text_frame
            for para in textf.paragraphs:
                if para.runs or para.text:
                    paragraphs.append(para)

    return paragraphs


# endregion

# region ===========
# endregion

# region docxtext2pptx pipeline
"""
Below are all the functions written for the docx2pptx original pipeline flow; we'll start the reverse flow above.
"""
# endregion


# region run_docxtext2pptx_pipeline()
def run_docx2pptx_pipeline(docx_path: Path) -> None:
    """Orchestrates the docx2pptx pipeline."""
    user_path = docx_path

    # Validate it's a real path of the correct type. If it's not, return the error.
    try:
        user_path_validated = validate_docx_path(user_path)
    except FileNotFoundError:
        print(f"Error: File not found: {user_path}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except PermissionError:
        print(f"I don't have permission to read that file ({user_path})!")
        sys.exit(1)

    # Load the docx file at that path.
    user_docx = open_and_load_docx(user_path_validated)

    # Chunk the docx by ___
    chunks = create_docx_chunks(user_docx, CHUNK_TYPE)

    if PRESERVE_DOCX_METADATA_IN_SPEAKER_NOTES:
        chunks = process_chunk_annotations(chunks, user_docx)

    # Create the presentation object from template
    try:
        output_prs = create_empty_slide_deck()
    except Exception as e:
        print(f"Could not load template file (may be corrupted): {e}")
        sys.exit(1)

    # Mutate the presentation object by adding slides
    slides_from_chunks(user_docx, output_prs, chunks)

    # Save the presentation to an actual pptx on disk

    save_output(output_prs)


# endregion

# region create_slides.py
# eventual destination: ./src/docx2pptx/create_slides.py


def _copy_basic_run_formatting(
    source_font: Union[Font_docx, Font_pptx], target_font: Union[Font_docx, Font_pptx]
) -> None:
    """Extract common formatting logic for Runs."""

    # Bold/Italics: Only overwrite when explicitly set on the source (avoid clobbering inheritance)
    if source_font.bold is not None:
        target_font.bold = source_font.bold
    if source_font.italic is not None:
        target_font.italic = source_font.italic

    # Underline: collapse any explicit value (True/False/WD_UNDERLINE.*) to bool
    if source_font.underline is not None:
        target_font.underline = bool(source_font.underline)

    # TODO: Test size
    if source_font.size is not None:
        target_font.size = Pt(source_font.size.pt)
        """
        <a:r>
            <a:rPr lang="en-US" sz="8800" i="1" dirty="0"/>
            <a:t>MAKE this text BIG!</a:t>
        </a:r>
        """


def _copy_run_color_formatting(
    source_font: Union[Font_docx, Font_pptx], target_font: Union[Font_docx, Font_pptx]
) -> None:
    # Color: copy only if source has an explicit RGB
    src_rgb = getattr(getattr(source_font, "color", None), "rgb", None)
    if src_rgb is not None:
        if isinstance(target_font, Font_pptx):
            target_font.color.rgb = RGBColor_pptx(*src_rgb)
        elif isinstance(target_font, Font_docx):
            target_font.color.rgb = RGBColor_docx(*src_rgb)


def copy_run_formatting_docx2pptx(
    source_run: Run_docx,
    target_run: Run_pptx,
    source_paragraph: Paragraph_docx,
    experimental_formatting_metadata: list,
) -> None:
    """Mutates a pptx _Run object to apply text and formatting from a docx Run object."""
    sfont = source_run.font
    tfont = target_run.font

    target_run.text = source_run.text

    _copy_basic_run_formatting(sfont, tfont)

    _copy_run_color_formatting(sfont, tfont)

    if isinstance(source_run, Run_docx) and EXPERIMENTAL_FORMATTING_ON:
        if source_run.text and source_run.text.strip():
            _copy_experimental_formatting_docx2pptx(
                source_run, target_run, experimental_formatting_metadata
            )

    # TODO: can we use the run object's "dirty" XML attribute to force-auto-resize slide text?


# TODO: Extracting all the experimental formatting for the purpose of restoring during the reverse pipeline.
def _copy_experimental_formatting_docx2pptx(
    source_run: Run_docx,
    target_run: Run_pptx,
    experimental_formatting_metadata: list,
) -> None:
    sfont = source_run.font
    tfont = target_run.font

    # The following code, which extends formatting support beyond python-pptx's capabilities,
    # is adapted from the md2pptx project, particularly from ./paragraph.py
    # Original source: https://github.com/MartinPacker/md2pptx
    # Author: Martin Packer
    # License: MIT
    if sfont.highlight_color is not None:
        # TODO: ADD TO THE experimental_formatting_metadata list
        experimental_formatting_metadata.append(
            {
                "ref_text": source_run.text,
                "highlight_color": COLOR_MAP_HEX.get(sfont.highlight_color),
                "formatting_type": "highlight",
            }
        )
        try:
            # Convert the docx run highlight color to a hex string
            tfont_hex_str = COLOR_MAP_HEX.get(sfont.highlight_color)

            # Create an object to represent this run in memory
            rPr = target_run._r.get_or_add_rPr()  # type: ignore[reportPrivateUsage]

            # Create a highlight Oxml object in memory
            hl = OxmlElement("a:highlight")

            # Create a srgbClr Oxml object in memory
            srgbClr = OxmlElement("a:srgbClr")

            # Set the attribute val of the srgbClr Oxml object in memory to the desired color
            setattr(srgbClr, "val", tfont_hex_str)

            # Add srgbClr object inside the hl Oxml object
            hl.append(srgbClr)  # type: ignore[reportPrivateUsage]

            # Add the hl object to the run representation object, which will add all our Oxml elements inside it
            rPr.append(hl)  # type: ignore[reportPrivateUsage]

        except Exception as e:
            debug_print(
                f"We found a highlight in the docx run but couldn't apply it. \n Run text: {source_run.text[:50]}... \n Error: {e}"
            )
            debug_print(
                "In order to attempt to preserve the visual difference of the highlighted text, we'll apply a basic gradient effect instead."
            )
            tfont.fill.gradient()
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

    if sfont.strike is not None:
        experimental_formatting_metadata.append(
            {"ref_text": source_run.text, "formatting_type": "strike"}
        )
        try:
            tfont._element.set("strike", "sngStrike")  # type: ignore[reportPrivateUsage]
        except Exception as e:
            debug_print(
                f"Failed to apply strikethrough. \nRun text: {source_run.text[:50]}... \n Error: {e}"
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

    if sfont.double_strike is not None:
        experimental_formatting_metadata.append(
            {"ref_text": source_run.text, "formatting_type": "double_strike"}
        )
        try:
            tfont._element.set("strike", "dblStrike")  # type: ignore[reportPrivateUsage]
        except Exception as e:
            debug_print(
                f"""
                        Failed to apply double-strikthrough.
                        \nRun text: {source_run.text[:50]}... \n Error: {e}
                        \nWe'll attempt single strikethrough."""
            )
            tfont._element.set("strike", "sngStrike")  # type: ignore[reportPrivateUsage]
        """
        Reference pptx XML for double strikethrough:
        <a:p>
            <a:r>
                <a:rPr lang="en-US" strike="dblStrike" dirty="0" err="1"/>
                <a:t>Double strike this text.</a:t>
            </a:r>        
        </a:p>
        """

    if sfont.subscript is not None:
        experimental_formatting_metadata.append(
            {"ref_text": source_run.text, "formatting_type": "subscript"}
        )
        try:
            if tfont.size is None:
                tfont._element.set("baseline", "-50000")  # type: ignore[reportPrivateUsage]

            if tfont.size is not None and tfont.size < Pt(24):
                tfont._element.set("baseline", "-50000")  # type: ignore[reportPrivateUsage]
            else:
                tfont._element.set("baseline", "-25000")  # type: ignore[reportPrivateUsage]

        except Exception as e:
            debug_print(
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

    if sfont.superscript is not None:
        experimental_formatting_metadata.append(
            {"ref_text": source_run.text, "formatting_type": "superscript"}
        )
        try:
            if tfont.size is None:
                tfont._element.set("baseline", "60000")  # type: ignore[reportPrivateUsage]

            if tfont.size is not None and tfont.size < Pt(24):
                tfont._element.set("baseline", "60000")  # type: ignore[reportPrivateUsage]
            else:
                tfont._element.set("baseline", "30000")  # type: ignore[reportPrivateUsage]

        except Exception as e:
            debug_print(
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
    if sfont.all_caps is not None:
        experimental_formatting_metadata.append(
            {"ref_text": source_run.text, "formatting_type": "all_caps"}
        )
        try:
            tfont._element.set("cap", "all")  # type: ignore[reportPrivateUsage]
        except Exception as e:
            debug_print(
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

    if sfont.small_caps is not None:
        experimental_formatting_metadata.append(
            {"ref_text": source_run.text, "formatting_type": "small_caps"}
        )
        try:
            tfont._element.set("cap", "small")  # type: ignore[reportPrivateUsage]
        except Exception as e:
            debug_print(
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


def create_blank_slide_for_chunk(
    prs: presentation.Presentation, slide_layout: SlideLayout
) -> tuple[Slide, TextFrame]:
    """Initialize an empty slide so that we can populate it with a chunk."""
    new_slide = prs.slides.add_slide(slide_layout)
    content = new_slide.placeholders[1]

    if not isinstance(content, SlidePlaceholder):
        raise TypeError(f"Expected SlidePlaceholder, got {type(content)}")

    text_frame: TextFrame = content.text_frame
    text_frame.clear()

    # Access the slide's notes_slide attribute in order to initialize it.
    notes_slide_ptr = new_slide.notes_slide  # type: ignore # noqa

    return new_slide, text_frame


def slides_from_chunks(
    doc: document.Document,
    prs: presentation.Presentation,
    chunks: list[Chunk_docx],
) -> None:
    """Generate slide objects, one for each chunk created by earlier pipeline steps."""

    # Specify which slide layout to use
    slide_layout = prs.slide_layouts.get_by_name(SLD_LAYOUT_CUSTOM_NAME)

    if slide_layout is None:
        raise KeyError(
            f"No slide layout found to match provided custom name, {SLD_LAYOUT_CUSTOM_NAME}"
        )

    for chunk in chunks:
        # Create a new slide for this chunk.
        new_slide, text_frame = create_blank_slide_for_chunk(prs, slide_layout)

        # Store custom metadata for this chunk that we'll want to tuck into the speaker notes as JSON
        # (for the purposes of restoring during reverse pipeline runs).
        slide_metadata = {}
        headings = []
        experimental_formatting: list[dict] = []

        # For each paragraph in this chunk, handle adding it
        for i, paragraph in enumerate(chunk.paragraphs):

            # Creating a new slide and a text frame leaves an empty paragraph in place, even when clearing it.
            # So if we're at the start of our list, use that existing empty paragraph.
            if (  # TODO: can we use this logic in the reverse pipeline flow for para 0 so it's not always empty?
                i == 0
                and len(text_frame.paragraphs) > 0
                and not text_frame.paragraphs[0].text
            ):
                # Use the existing first/0th paragraph
                pptx_paragraph = text_frame.paragraphs[0]
            else:
                pptx_paragraph = text_frame.add_paragraph()

            # Process the docx's paragraph contents, including both runs & hyperlinks
            para_experimental_formatting = process_chunk_paragraph_inner_contents(
                paragraph, pptx_paragraph
            )

            if para_experimental_formatting:
                experimental_formatting.extend(para_experimental_formatting)

            if paragraph.style and is_standard_heading(paragraph.style.style_id):
                headings.append(
                    {
                        "text": paragraph.text.strip(),
                        "style_id": paragraph.style.style_id,
                    }
                )

        if headings:
            slide_metadata["headings"] = headings
        if experimental_formatting:
            slide_metadata["experimental_formatting"] = experimental_formatting

        notes_text_frame: TextFrame = new_slide.notes_slide.notes_text_frame  # type: ignore # TODO, POLISH: is there a reasonable way to fix this type hint?

        if DISPLAY_DOCX_ANNOTATIONS_IN_SLIDE_SPEAKER_NOTES:
            annotate_slide(chunk, notes_text_frame)

        if PRESERVE_DOCX_METADATA_IN_SPEAKER_NOTES:
            add_metadata_to_slide_notes(notes_text_frame, chunk, slide_metadata)


def add_metadata_to_slide_notes(
    notes_text_frame: TextFrame, chunk: Chunk_docx, slide_body_metadata: dict
) -> None:
    """
    Populate the slide notes text frame with docx metadata so that we may restore it during a round-trip pptx2docx pipeline.
    """
    annotation_metadata = {}

    if chunk.comments:
        annotation_metadata["comments"] = [
            {
                "original": {
                    "text": c.comment_obj.text,
                    "author": c.comment_obj.author,
                    "comment_id": c.comment_obj.comment_id,
                    "initials": c.comment_obj.initials,
                    "timestamp": (
                        str(c.comment_obj.timestamp)
                        if c.comment_obj.timestamp
                        else None
                    ),
                },
                "reference_text": c.reference_text,
                "id": c.note_id,
            }
            for c in chunk.comments
        ]

    if chunk.footnotes:
        annotation_metadata["footnotes"] = [
            {
                "footnote_id": f.footnote_id,
                "text_body": f.text_body,
                "hyperlinks": f.hyperlinks,
                "reference_text": f.reference_text,
            }
            for f in chunk.footnotes
        ]

    if chunk.endnotes:
        annotation_metadata["endnotes"] = [
            {
                "endnote_id": e.endnote_id,
                "text_body": e.text_body,
                "hyperlinks": e.hyperlinks,
                "reference_text": e.reference_text,
            }
            for e in chunk.endnotes
        ]

    if slide_body_metadata or annotation_metadata:
        header_para = notes_text_frame.add_paragraph()
        header_run = header_para.add_run()
        header_run.text = (
            "\n\n\n\n\n\n\nSTART OF JSON METADATA FROM SOURCE DOCUMENT:\n" + "=" * 40
        )

        json_para = notes_text_frame.add_paragraph()
        json_run = json_para.add_run()

        combined_metadata = {**annotation_metadata, **slide_body_metadata}
        json_run.text = json.dumps(combined_metadata, indent=2)

        footer_para = notes_text_frame.add_paragraph()
        footer_run = footer_para.add_run()
        footer_run.text = "=" * 40 + "\nEND OF JSON METADATA FROM SOURCE DOCUMENT"


def process_chunk_paragraph_inner_contents(
    paragraph: Paragraph_docx, pptx_paragraph: Paragraph_pptx
) -> list[dict]:
    """
    Iterate through a paragraph's runs and hyperlinks, in document order, and:
    - copy it into the slide, with formatting
    - capture any additional metadata (like experimental formatting) that we cannot apply directly to the copied-over
        pptx objects
    """
    items_processed = False
    experimental_formatting_metadata = []

    for item in paragraph.iter_inner_content():
        items_processed = True
        if isinstance(item, Run_docx):

            # If this Run has a field code for instrText and it begins with HYPERLINK, this is an old-style
            # word hyperlink, which we cannot handle the same way as normal docx hyperlinks. But we try to detect
            # when it happens and report it to the user.
            if "instrText" in item.element.xml and "HYPERLINK" in item.element.xml:
                detect_field_code_hyperlinks(item)

            process_run(
                item, paragraph, pptx_paragraph, experimental_formatting_metadata
            )

        elif hasattr(item, "url"):
            # Process all runs within the hyperlink
            for run in item.runs:
                process_run(
                    run,
                    paragraph,
                    pptx_paragraph,
                    experimental_formatting_metadata,
                    item.url,
                )
        else:
            debug_print(f"Unknown content type in paragraph: {type(item)}")

    # Fallback: if no content was processed but paragraph has text
    if not items_processed and paragraph.text:
        debug_print(
            f"Fallback: paragraph has text but no runs/hyperlinks: {paragraph.text[:50]}"
        )
        pptx_run = pptx_paragraph.add_run()
        pptx_run.text = paragraph.text

    return experimental_formatting_metadata


def detect_field_code_hyperlinks(run: Run_docx) -> None:
    """
    Detect if this Run has a field code for instrText and it begins with HYPERLINK.
    If so, report it to the user, because we do not handle adding these to the pptx output.
    """
    try:
        run_xml: str = run.element.xml  # type: ignore
        root = ET.fromstring(run_xml)

        # Find instrText elements
        instr_texts = root.findall(
            ".//w:instrText",
            {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        )
        for instr in instr_texts:
            if instr.text and instr.text.startswith("HYPERLINK"):
                # TODO, polish, leafy: Add a const switch that would allow us to simply add the string {instr.text}
                # into the main text body if the user desires
                debug_print(
                    f"WARNING: We found a field code hyperlink, but we don't have a way to attach it to any text: {instr.text}"
                )

    except (AttributeError, ET.ParseError) as e:
        debug_print(
            f"WARNING: Could not parse run XML for field codes: {e} while seeking instrText"
        )


def process_run(
    run: Run_docx,
    docx_paragraph: Paragraph_docx,
    pptx_paragraph: Paragraph_pptx,
    experimental_formatting_metadata: list,
    hyperlink: str | None = None,
) -> Run_pptx:
    """Copy a run from the docx parent to the pptx paragraph, including copying its formatting."""
    # Handle formatting

    pptx_run = pptx_paragraph.add_run()
    copy_run_formatting_docx2pptx(
        run, pptx_run, docx_paragraph, experimental_formatting_metadata
    )

    if hyperlink:
        pptx_run_url = pptx_run.hyperlink
        pptx_run_url.address = hyperlink

    return pptx_run


# endregion

# region GET annotation helpers
# eventual destination: ./src/docx2pptx/annotations/annotate_chunks.py
# endregion


# region annotate_chunks.py
def process_chunk_annotations(
    chunks: list[Chunk_docx], doc: document.Document
) -> list[Chunk_docx]:
    """For a list of Chunk_docx objects, populate the annotation dicts for each one."""

    # Gather all the doc annotations
    all_raw_comments = get_all_docx_comments(doc)
    all_footnotes = get_all_docx_footnotes(doc)
    all_endnotes = get_all_docx_endnotes(doc)

    for chunk in chunks:
        for paragraph in chunk.paragraphs:
            for item in paragraph.iter_inner_content():

                if isinstance(item, Run_docx):
                    process_run_annotations(
                        chunk,
                        paragraph,
                        item,
                        all_raw_comments=all_raw_comments,
                        all_footnotes=all_footnotes,
                        all_endnotes=all_endnotes,
                    )

                # If the item has the attribute "url" we assume it is of type Hyperlink instead of Run;
                # that means it contains its own child runs, so we need to go one step inward to process them.
                elif hasattr(item, "url"):
                    # Process all runs within the hyperlink
                    for run in item.runs:
                        process_run_annotations(
                            chunk,
                            paragraph,
                            run,
                            all_raw_comments=all_raw_comments,
                            all_footnotes=all_footnotes,
                            all_endnotes=all_endnotes,
                        )
                else:
                    debug_print(f"Unknown content type in paragraph: {type(item)}")

    return chunks


def process_run_annotations(
    chunk: Chunk_docx,
    paragraph: Paragraph_docx,
    run: Run_docx,
    all_raw_comments: dict[str, Comment_docx],
    all_footnotes: dict[str, Footnote_docx],
    all_endnotes: dict[str, Endnote_docx],
) -> None:
    """Get the annotations from a run object and adding them into its (grand)parent chunk object."""
    try:
        # Get XML from the run using public API
        run_xml = run.element.xml  # type: ignore

        # Parse it safely with ElementTree
        root = ET.fromstring(run_xml)

        # Define namespace
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # get the reference text to be used by comments, footnotes, or endnotes
        ref_text = get_ref_text(run, paragraph)

        # Find comment references
        comment_refs = root.findall(".//w:commentReference", ns)
        for ref in comment_refs:
            comment_id = ref.get(f'{{{ns["w"]}}}id')
            if comment_id and comment_id in all_raw_comments:
                comment_object = all_raw_comments[comment_id]

                custom_comment_obj = Comment_docx_custom(
                    comment_obj=comment_object, reference_text=ref_text
                )
                chunk.add_comment(custom_comment_obj)

        # Find footnote references
        footnote_refs = root.findall(".//w:footnoteReference", ns)
        for ref in footnote_refs:
            footnote_id = ref.get(f'{{{ns["w"]}}}id')
            if footnote_id and footnote_id in all_footnotes:
                footnote_obj = all_footnotes[footnote_id]
                footnote_obj.reference_text = ref_text
                chunk.add_footnote(footnote_obj)

        # Find endnote references - same pattern
        endnote_refs = root.findall(".//w:endnoteReference", ns)
        for ref in endnote_refs:
            endnote_id = ref.get(f'{{{ns["w"]}}}id')
            if endnote_id and endnote_id in all_endnotes:
                endnote_obj = all_endnotes[endnote_id]
                endnote_obj.reference_text = ref_text
                chunk.add_endnote(endnote_obj)

    except (AttributeError, ET.ParseError) as e:
        debug_print(f"WARNING: Could not parse run XML for references: {e}")


# endregion


def get_ref_text(run: Run_docx, paragraph: Paragraph_docx) -> str | None:
    """
    Get the Run or Paragraph text with which a piece of metadata is associated in the docx so that we can store that in
    metadata and reference it on reverse-pipeline runs.
    """
    if run.text and run.text.strip():
        ref_text = run.text
    elif paragraph.text and paragraph.text.strip():
        # Grab the first (up to 10) words of this paragraph if the run text is empty
        ref_text = " ".join(paragraph.text.split()[:10])
    else:
        ref_text = None

    return ref_text


# eventual destination: ./src/docx2pptx/annotations/get_docx_annotations.py
# region Get all notes in this docx
def get_all_docx_comments(doc: document.Document) -> dict[str, Comment_docx]:
    """
    Get all the comments in this document as a dictionary.

    Elements of the dictionary are formatted as:
    {
        "comment_id_#" : this_comment_docx_object,
        "3": "<docx.comments.Comment object at 0x00000###>
    }
    """
    all_comments_dict: dict[str, Comment_docx] = {}

    if hasattr(doc, "comments") and doc.comments:
        for comment in doc.comments:
            all_comments_dict[str(comment.comment_id)] = comment
    return all_comments_dict


def get_all_docx_footnotes(doc: document.Document) -> dict[str, Footnote_docx]:
    """
    Extract all footnotes from a docx document.
    Returns {id: {footnote_id: str, text_body: str, hyperlinks: list of str} }.
    """

    if not DISPLAY_FOOTNOTES:
        return {}

    try:
        footnote_parts = find_xml_parts(doc, "footnotes.xml")

        if not footnote_parts:
            return {}

        # We think this will always be a list of one item, so assign that item to a variable.
        root = parse_xml_blob(footnote_parts[0].blob)
        return extract_notes_from_xml(root, Footnote_docx)

    except Exception as e:
        debug_print(f"Warning: Could not extract footnotes: {e}")
        return {}


def get_all_docx_endnotes(doc: document.Document) -> dict[str, Endnote_docx]:
    """
    Extract all endnotes from a docx document.
    Returns {id: {footnote_id: str, text_body: str, hyperlinks: list of str} }.
    """
    if not DISPLAY_ENDNOTES:
        return {}

    try:
        endnote_parts = find_xml_parts(doc, "endnotes.xml")

        if not endnote_parts:
            return {}

        root = parse_xml_blob(endnote_parts[0].blob)
        return extract_notes_from_xml(root, Endnote_docx)

    except Exception as e:
        debug_print(f"Warning: Could not extract endnotes: {e}")
        return {}


# endregion


# region Utils - XML parsing
def find_xml_parts(doc: document.Document, part_name: str) -> list[Part]:
    """Find XML parts matching the given name (e.g., 'footnotes.xml')"""
    # The zip package inspection logic
    # Inspect the docx package as a zip
    zip_package = doc.part.package

    if zip_package is None:
        debug_print("WARNING: Could not access docx package")
        return []

    part_name_parts: list[Part] = []
    for part in zip_package.parts:
        if part_name in str(part.partname):
            debug_print(f"We found a {part_name} part!")
            part_name_parts.append(part)

    return part_name_parts


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


# TypeVar definition; TODO: move to the top of whatever file this function ends up living in
# Generic type parameter - when you pass Footnote_docx into the below function, you will get dict[str, Footnote_docx] back
NOTE_TYPE = TypeVar("NOTE_TYPE", Footnote_docx, Endnote_docx)


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
        note_hyperlinks = extract_hyperlinks_from_note(note)

        note_obj = note_class(
            note_id, text_body=note_full_text, hyperlinks=note_hyperlinks
        )

        notes_dict[note_id] = note_obj

    return notes_dict


def extract_hyperlinks_from_note(element: ET.Element) -> list[str]:
    """Extract all hyperlinks from a footnote element."""
    hyperlinks: list[str] = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    for hyperlink in element.findall(".//w:hyperlink", ns):
        # Get the link text
        link_text = "".join([t.text or "" for t in hyperlink.findall(".//w:t", ns)])
        if link_text.strip():
            hyperlinks.append(link_text.strip())

    return hyperlinks


# endregion


# region annotate_slides.py
# eventual destination: ./src/docx2pptx/annotations/annotate_slides.py


def annotate_slide(chunk: Chunk_docx, notes_text_frame: TextFrame) -> None:
    """
    Pull a chunk's preserved annotations and copy them into the slide's speaker notes text frame.

    NOTE: We DO NOT PRESERVE any anchoring to the slide's body for annotations. That means we don't preserve
    comments' selected text ranges, nor do we preserve footnote or endnote numbering.
    """
    if DISPLAY_COMMENTS and chunk.comments:
        add_comments_to_notes(chunk.comments, notes_text_frame)

    if DISPLAY_FOOTNOTES and chunk.footnotes:
        add_footnotes_to_notes(chunk.footnotes, notes_text_frame)

    if DISPLAY_ENDNOTES and chunk.endnotes:
        add_endnotes_to_notes(chunk.endnotes, notes_text_frame)

    if PRESERVE_DOCX_METADATA_IN_SPEAKER_NOTES:
        pass
        # add_metadata_to_slide_notes(chunk.metadata, notes_text_frame)


# region Add annotations to pptx speaker notes text frame
def add_comments_to_notes(
    comments_list: list[Comment_docx_custom], notes_text_frame: TextFrame
) -> None:
    """Copy logic for appending the comments portion of the speaker notes."""

    if comments_list:
        if COMMENTS_SORT_BY_DATE:
            # Sort comments by date (newest first, or change reverse=False for oldest first)
            sorted_comments = sorted(
                comments_list,
                key=lambda c: getattr(c.comment_obj, "timestamp", None) or datetime.min,
                reverse=False,
            )
        else:
            sorted_comments = comments_list

        comment_para = notes_text_frame.add_paragraph()
        comment_run = comment_para.add_run()
        comment_run.text = "COMMENTS FROM SOURCE DOCUMENT:\n" + "=" * 40

        for i, comment in enumerate(sorted_comments, 1):
            # Check if comment has paragraphs attribute
            if hasattr(comment.comment_obj, "paragraphs"):
                # Get paragraphs safely, default to empty list if not present
                this_comment_paragraphs = getattr(comment.comment_obj, "paragraphs", [])
                for para in this_comment_paragraphs:
                    if hasattr(para, "text") and para.text.rstrip():
                        notes_para = notes_text_frame.add_paragraph()
                        comment_header = notes_para.add_run()

                        if COMMENTS_KEEP_AUTHOR_AND_DATE:
                            author = getattr(
                                comment.comment_obj, "author", "Unknown Author"
                            )
                            timestamp = getattr(comment.comment_obj, "timestamp", None)

                            if timestamp and hasattr(timestamp, "strftime"):
                                timestamp_str = timestamp.strftime(
                                    "%A, %B %d, %Y at %I:%M %p"
                                )
                            else:
                                timestamp_str = "Unknown Date"

                            comment_header.text = (
                                f"\n {i}. {author} ({timestamp_str}):\n"
                            )

                        else:
                            comment_header.text = "\n"
                        process_chunk_paragraph_inner_contents(para, notes_para)


# EXPERIMENT
def add_notes_to_speaker_notes(
    notes_list: list[Footnote_docx] | list[Endnote_docx],
    notes_text_frame: TextFrame,
    note_class: type[NOTE_TYPE],
) -> None:
    """Generic function for adding footnotes or endnotes to speaker notes."""

    if note_class is Footnote_docx:
        header_text = "FOOTNOTES FROM SOURCE DOCUMENT"
    elif note_class is Endnote_docx:
        header_text = "ENDNOTES FROM SOURCE DOCUMENT"
    else:
        header_text = "Unknown Note Type from Source Document:"

    if notes_list:
        note_para = notes_text_frame.add_paragraph()
        note_run = note_para.add_run()
        note_run.text = f"\n{header_text}:\n" + "=" * 40

        for note_obj in notes_list:
            notes_para = notes_text_frame.add_paragraph()
            note_run = notes_para.add_run()

            # Start with the main note text
            note_text = f"\n{note_obj.note_id}. {note_obj.text_body}\n"

            # Add hyperlinks if they exist
            if note_obj.hyperlinks:
                note_text += "\nHyperlinks:"
                for hyperlink in note_obj.hyperlinks:
                    note_text += f"\n{hyperlink}"

            note_run.text = note_text


# TODO, ANNOTATIONS: Replace specific functions with this generic one, and test
def add_footnotes_to_notes(
    footnotes_list: list[Footnote_docx], notes_text_frame: TextFrame
) -> None:
    """Copy logic for appending the footnotes portion of the speaker notes."""

    if footnotes_list:
        footnote_para = notes_text_frame.add_paragraph()
        footnote_run = footnote_para.add_run()
        footnote_run.text = "\nFOOTNOTES FROM SOURCE DOCUMENT:\n" + "=" * 40

        for footnote_obj in footnotes_list:
            notes_para = notes_text_frame.add_paragraph()
            footnote_run = notes_para.add_run()
            footnote_text = f"\n{footnote_obj.footnote_id}. {footnote_obj.text_body}\n"

            # Add any hyperlinks if they exist
            if footnote_obj.hyperlinks:
                footnote_text += "\nHyperlinks:"
                for hyperlink in footnote_obj.hyperlinks:
                    footnote_text += f"\n{hyperlink}"

            # Assign the complete text to the run
            footnote_run.text = footnote_text


def add_endnotes_to_notes(
    endnotes_list: list[Endnote_docx], notes_text_frame: TextFrame
) -> None:
    """Copy logic for appending the endnote portion of the speaker notes."""

    if endnotes_list:
        endnote_para = notes_text_frame.add_paragraph()
        endnote_run = endnote_para.add_run()
        endnote_run.text = "\nENDNOTES FROM SOURCE DOCUMENT:\n" + "=" * 40

        for endnote_obj in endnotes_list:
            notes_para = notes_text_frame.add_paragraph()
            endnote_run = notes_para.add_run()
            # Start with the main endnote text
            endnote_text = f"\n{endnote_obj.endnote_id}. {endnote_obj.text_body}\n"

            # Add any hyperlinks if they exist
            if endnote_obj.hyperlinks:
                endnote_text += "\nHyperlinks:"
                for hyperlink in endnote_obj.hyperlinks:
                    endnote_text += f"\n{hyperlink}"

            # Assign the complete text to the run
            endnote_run.text = endnote_text


# endregion


# region create_chunks.py
# eventual destination: ./src/docx2pptx/create_chunks.py
def create_docx_chunks(
    doc: document.Document, chunk_type: ChunkType = ChunkType.PARAGRAPH
) -> list[Chunk_docx]:
    """
    Orchestrator function to create chunks (that will become slides) from the document
    contents, either from paragraph, heading (heading_nested or heading_flat),
    or page. Defaults to paragraph.
    """
    if chunk_type == ChunkType.HEADING_FLAT:
        chunks = chunk_by_heading_flat(doc)
    elif chunk_type == ChunkType.HEADING_NESTED:
        chunks = chunk_by_heading_nested(doc)
    elif chunk_type == ChunkType.PAGE:
        chunks = chunk_by_page(doc)
    else:
        chunks = chunk_by_paragraph(doc)
    return chunks


# endregion


# region by Paragraph
def chunk_by_paragraph(doc: document.Document) -> list[Chunk_docx]:
    """
    Creates chunks (which will become slides) based on paragraph, which are blocks of content
    separated by whitespace.
    """
    paragraph_chunks: list[Chunk_docx] = []

    for para in doc.paragraphs:

        # Skip empty paragraphs (but keep those that are new-lines to respect intentional whitespace newlines)
        if para.text == "":
            continue

        new_chunk = Chunk_docx.create_with_paragraph(para)
        paragraph_chunks.append(new_chunk)

    return paragraph_chunks


# endregion


# region by Page
def chunk_by_page(doc: document.Document) -> list[Chunk_docx]:
    """Creates chunks based on page breaks"""

    # Start building the chunks
    all_chunks: list[Chunk_docx] = []

    # Start with a current chunk ready-to-go
    current_page_chunk: Chunk_docx = Chunk_docx()

    for para in doc.paragraphs:
        # Skip empty paragraphs (keep intentional whitespace newlines)
        if para.text == "":
            continue

        # If the current_page_chunk is empty, append the current para regardless of style & continue to next para.
        if not current_page_chunk.paragraphs:
            current_page_chunk.add_paragraph(para)
            continue

        # Handle page breaks - create new chunk and start fresh
        if para.contains_page_break:
            # Add the current_chunk to chunks list (if it has content)
            if current_page_chunk:
                all_chunks.append(current_page_chunk)

            # Start new chunk with this paragraph
            current_page_chunk = Chunk_docx.create_with_paragraph(para)

            continue

        # If there was no page break, just append this paragraph to the current_chunk
        current_page_chunk.add_paragraph(para)

    # Ensure final chunk from loop is added to chunks list
    if current_page_chunk:
        all_chunks.append(current_page_chunk)

    print(f"This document has {len(all_chunks)} page chunks.")
    return all_chunks


# endregion

# region by Heading (nested)


## === chunk by heading helpers ===


def is_standard_heading(style_id: str) -> bool:
    """Check if style_id is a standard Word Heading (Heading1, Heading2, ..., Heading6)"""
    return style_id.startswith("Heading") and style_id[7:].isdigit()


def get_heading_level(style_id: str) -> int | float:
    """
    Extract the numeric level from a heading style (e.g., 'Heading2' -> 2),
    or return infinity if the style_id doesn't have a number.
    """
    try:
        return int(style_id[7:])
    except (ValueError, IndexError):
        return float("inf")  # Treat non-headings as "deepest possible"


## ===


def chunk_by_heading_nested(doc: document.Document) -> list[Chunk_docx]:
    """
    Creates chunks based on headings, using nesting logic to group "deeper" headings

    New chunks are begun when:
    - a page break happens in the middle of a paragraph
    - we reach a heading-depth that is equal to or "higher" than (number is less than) the current depth
    Otherwise, appends paragraph to the current chunk.

    E.g.:

    CHUNK:
    H1
    Normal Paragraph
    H2
    Normal Paragraph
    Normal Paragraph

    NEXT_CHUNK:
    H2
    Normal Paragraph
    Normal Paragraph
    Normal Paragraph

    NEXT_CHUNK:
    H1
    H2
    Pararaph
    Normal Paragraph
    H3
    Normal Paragraph

    NEXT_CHUNK:
    H2
    Normal Paragraph
    Normal Paragraph

    """
    # Start building the chunks
    all_chunks: list[Chunk_docx] = []
    current_chunk: Chunk_docx = Chunk_docx()

    # Initialize current_heading_style_id
    current_heading_style_id = "Normal"  # Default for documents without headings

    for i, para in enumerate(doc.paragraphs):

        # Skip empty paragraphs
        if para.text == "":
            continue

        # Set a style_id to make Pylance happy (it gets mad if we direct-check para.style.style_id later)
        style_id = para.style.style_id if para.style else "Normal"

        debug_print(f"Paragraph begins: {para.text[:30]}... and is index: {i}")

        # If the current_chunk is empty, append the current para regardless of style & continue to next para.
        if not current_chunk.paragraphs:
            current_chunk.add_paragraph(para)
            if is_standard_heading(style_id):
                # TODO: store this heading information somewhere
                current_heading_style_id = style_id
            continue

        # Handle page breaks - create new chunk and start fresh
        if para.contains_page_break:
            # Add the current chunk to chunks list (if it has content)
            if current_chunk:
                all_chunks.append(current_chunk)

            # Start new chunk with this paragraph
            current_chunk = Chunk_docx.create_with_paragraph(para)

            # Update heading depth if this paragraph is a heading
            if is_standard_heading(style_id):
                # TODO: store this heading information somewhere
                current_heading_style_id = style_id
            continue

        # Handle headings
        if is_standard_heading(style_id):
            # TODO: store this heading information somewhere
            # Check if this heading is at same level or higher (less deep) than current. Smaller numbers are higher up in the hierarchy.
            if get_heading_level(style_id) <= get_heading_level(
                current_heading_style_id
            ):
                # If yes, start a new chunk
                if current_chunk:
                    all_chunks.append(current_chunk)
                current_chunk = Chunk_docx.create_with_paragraph(para)
                current_heading_style_id = style_id
            else:
                # This heading is deeper, add to current chunk
                current_chunk.add_paragraph(para)
                current_heading_style_id = style_id
        else:
            # Normal paragraph - add to current chunk
            current_chunk.add_paragraph(para)

    if current_chunk:
        all_chunks.append(current_chunk)

    print(f"This document has {len(all_chunks)} nested heading chunks.")
    return all_chunks


# endregion


# region by Heading (flat)
def chunk_by_heading_flat(doc: document.Document) -> list[Chunk_docx]:
    """
    Creates chunks based on headings; no nesting logic used. New chunks are created when:
    - a page break happens in the middle of a paragraph
    - we reach any paragraph that is any heading

    CHUNK:
    H1
    Normal Paragraph

    NEXT_CHUNK:
    H2
    Normal Paragraph
    Normal Paragraph

    NEXT_CHUNK:
    H2
    Normal Paragraph
    Normal Paragraph
    Normal Paragraph

    NEXT_CHUNK:
    H1

    NEXT_CHUNK:
    H2
    Pararaph
    Normal Paragraph

    NEXT_CHUNK:
    H3
    Normal Paragraph

    NEXT_CHUNK:
    H2
    Normal Paragraph
    Normal Paragraph
    """

    # Start building the chunks
    all_chunks: list[Chunk_docx] = []
    current_chunk: Chunk_docx = Chunk_docx()

    for para in doc.paragraphs:
        # Skip empty paragraphs
        if para.text == "":
            continue

        # Set a style_id to make Pylance happy (it gets mad if we direct-check para.style.style_id later)
        style_id = para.style.style_id if para.style else "Normal"

        debug_print(f"Paragraph begins: {para.text[:30]}...")

        # If the current_chunk is empty, append the current para regardless of style & continue to next para.
        if not current_chunk.paragraphs:
            current_chunk.add_paragraph(para)
            continue

        # Handle page breaks - always start a new chunk
        if para.contains_page_break:
            # Add the current chunk to chunks list (if it has content)
            if current_chunk:
                all_chunks.append(current_chunk)

            # Start new chunk with this paragraph
            current_chunk = Chunk_docx.create_with_paragraph(para)
            continue

        # If this paragraph is a heading, start a new chunk
        if is_standard_heading(style_id):
            # TODO: store this heading information somewhere
            # If we already have content in current_chunk, save it and start fresh
            if current_chunk:
                all_chunks.append(current_chunk)

            # Start new chunk with this paragraph
            current_chunk = Chunk_docx.create_with_paragraph(para)

        else:
            # This is a normal paragraph - add it to current chunk
            current_chunk.add_paragraph(para)

    if current_chunk:
        all_chunks.append(current_chunk)

    print(f"This document has {len(all_chunks)} flat heading chunks.")
    return all_chunks


# endregion


# region Utils - Basic
def debug_print(msg: str | list[str]) -> None:
    """Basic debug printing function"""
    if DEBUG_MODE:
        print(f"DEBUG: {msg}")


def setup_console_encoding() -> None:
    """Configure UTF-8 encoding for Windows console to prevent UnicodeEncodeError when printing non-ASCII characters (like emojis)."""
    if platform.system() == "Windows":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


# endregion


# region Utils - I/O
def validate_path(user_path: str | Path) -> Path:
    """Ensure filepath exists and is a file."""
    path = Path(user_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {user_path}")
    if not path.is_file():
        raise ValueError("That's not a file.")
    return path


def validate_docx_path(user_path: str | Path) -> Path:
    """Validates the user-provided filepath exists and is actually a docx file."""
    path = validate_path(user_path)
    # Verify it's the right extension:
    if path.suffix.lower() == ".doc":
        raise ValueError(
            "This tool only supports .docx files right now. Please convert your .doc file to .docx format first."
        )
    elif path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file, but got: {path.suffix}")

    # Add document structure validation
    try:
        doc = docx.Document(path)  # type: ignore
        if not doc.paragraphs:
            raise ValueError("Document appears to be empty")
        return path
    except Exception as e:
        raise ValueError(f"Document appears to be corrupted: {e}")


def validate_pptx_path(user_path: str | Path) -> Path:
    """Validates the pptx template filepath exists and is actually a pptx file."""
    path = validate_path(user_path)
    # Verify it's the right extension:
    if path.suffix.lower() != ".pptx":
        raise ValueError(f"Expected a .pptx file, but got: {path.suffix}")
    return path


# region io.py
# TODO, BASIC VALIDATION: Add basic validation for docx contents (copy the validation we did in CSharp)
# TODO 1_ validate the document body is not null
# TODO 2_ validate that there is at least 1 paragraph with text in it
# eventual destination: ./src/docx2pptx/io.py (?)
def open_and_load_docx(input_filepath: Path) -> document.Document:
    """Use python-docx to read in the docx file contents and store to a runtime variable."""
    doc = docx.Document(input_filepath)  # type: ignore

    # Count and report paragraphs to validate that we can see content in the file.
    paragraph_count = len(doc.paragraphs)
    print(f"This document has {paragraph_count} paragraphs in it.")
    if paragraph_count > 0:
        text = doc.paragraphs[0].text
        preview = text[:20] + ("..." if len(text) > 20 else "")
        print(f"The first paragraph's text is: {preview}")
    return doc


# TODO, BASIC VALIDATION: Add validation to ensure the template is in good shape
# Like make sure it has all the pieces we're going to rely on; slide masters and slide layout, etc.,
# and make sure whatever the slide layout name the user provided and wishes to use exists.
def create_empty_slide_deck() -> presentation.Presentation:
    """Load the PowerPoint template and create a new presentation object."""

    try:
        template_path = validate_pptx_path(Path(TEMPLATE_PPTX))
        prs = pptx.Presentation(str(template_path))
        return prs
    except Exception as e:
        raise ValueError(f"Could not load template file (may be corrupted): {e}")

    # === testing
    # slide_layout = prs.slide_layouts[SLD_LAYOUT_CUSTOM]
    # slide = prs.slides.add_slide(slide_layout)
    # content = slide.placeholders[1]
    # content.text = "Test Slide!" # type:ignore


# TODO, BASIC VALIDATION:: Add some kind of validation that we're not saving something that's like over 100 MB. Maybe set a const
# TODO, UX: Probably add some kind of file rotation so the last 5 or so outputs are preserved
OUTPUT_TYPE = TypeVar("OUTPUT_TYPE", document.Document, presentation.Presentation)


def save_output(save_object: OUTPUT_TYPE) -> None:
    """Save the generated output object to disk as a file. Genericized to output either docx or pptx depending on which pipeline is running."""
    if isinstance(save_object, document.Document):
        save_folder = OUTPUT_DOCX_FOLDER
        save_filename = OUTPUT_DOCX_FILENAME
    elif isinstance(save_object, presentation.Presentation):
        save_folder = OUTPUT_PPTX_FOLDER
        save_filename = OUTPUT_PPTX_FILENAME
    else:
        raise RuntimeError(f"Unexpected output object type: {save_object}")

    try:
        # Construct output path
        if save_folder:
            folder = Path(save_folder)
            folder.mkdir(parents=True, exist_ok=True)
            output_filepath = folder / save_filename
        else:
            output_filepath = Path(save_filename)
        save_object.save(str(output_filepath))
        print(f"Successfully saved to {output_filepath}")
    except PermissionError:
        raise PermissionError("Save failed: File may be open in another program")
    except Exception as e:
        raise RuntimeError(f"Save failed with error: {e}")


def sanitize_xml_text(text: str) -> str:
    """Remove characters that aren't valid in XML."""
    if not text:
        return ""

    # Remove NULL bytes and control characters (except tab, newline, carriage return)
    sanitized = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)

    # Ensure it's a proper string
    return str(sanitized)


# endregion

# region call main
if __name__ == "__main__":
    main()
# endregion
