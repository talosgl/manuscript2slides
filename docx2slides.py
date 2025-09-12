# pyright: strict
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
- heading_flat: New slide for each heading (any level)
- heading_nested: New slide based on heading hierarchy
- page: New slide for each page break

Example:
    python docx2slides.py

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
from pptx.dml.color import RGBColor
from pptx.slide import Slide, SlideLayout
from pptx.text.text import TextFrame, _Paragraph as Paragraph_pptx, _Run as Run_pptx  # type: ignore
from pptx.shapes.placeholder import SlidePlaceholder
import xml.etree.ElementTree as ET

# endregion

# region Overarching TODOs
"""
Must-Implement v0 Features:
- Complete feature to preserve docx footnotes & endnotes and append after comments in speaker notes.
- Reverse flow: export slide text frame content to docx paragraphs with the annotations kept as comments, inline
    - This could allow a flow where you can iterate back and forth (and remove the need to "update" an existing deck with manuscript updates)

v1 features I'd like:
- Create Documents/docx2pptx/input/output/resources structure
    - Copies sample files from app resources to user folders
    - Cleanup mode for debug runs
- Change consts configuration to use a class or something
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
    -   We do not support .doc, only .docx. If you have a .doc file, convert it to .docx using Word, Google Docs, 
        or LibreOffice before processing.

    -   We do not support .ppt, only .pptx.

    -   Field code hyperlinks not supported - Some hyperlinks (like the sample_doc.docx's first "Where are Data?" link) 
        are stored using Word's field code format and display as plain text instead of clickable links. The exact 
        conditions that cause this format are unclear, but it may occur with hyperlinks in headings or certain copy/paste 
        scenarios. We think most normal hyperlinks will work fine.

    -   Auto-fit text resizing doesn't work. PowerPoint only applies auto-fit when opened in the UI. 
        You can get around this manually with these steps:
            1. Open up the output presentation in PowerPoint Desktop > View > Slide Master
            2. Select the text frame object, right-click > Format Shape
            3. Click the Size & Properties icon {TODO ADD IMAGES}
            4. Click Text Box to see the options
            5. Toggle "Do not Autofit" and then back to "Shrink Text on overflow"
            6. Close Master View
            7. Now all the slides should have their text properly resized.

    -   ANNOTATIONS LIMITATIONS
        -   We collapse all comments, footnotes, and endnotes into a slide's speaker notes. PowerPoint itself doesn't 
            support real footnotes or endnotes at all. It does have a comments functionality, but the library used here 
            (python-pptx) doesn't support adding comments to slides yet. 

        -   Note that inline reference numbers (1, 2, 3, etc.) from the docx body are not preserved in the slide text - 
            only the annotation content appears in speaker notes.

        -   You can choose to preserve some comment metadata (author, timestamps) in plain text, but not threading.

        -   Footnotes and endnotes don't preserve formatting, only plain text.
    
    -   REVERSE FLOW LIMITATIONS
        -   The reverse flow (pptx2docx) is significantly less robust for annotations. There is no support for separating out
            speaker notes into types of annotations. "We simply take all speaker notes and attach them to the final paragraph 
            of each slide's content as a comment. Original source docx comment metadata is not preserved (apart from
            plain text author/timestamps if kept during an original docx2pptx conversion).

"""
# endregion

# region CONSTANTS / config.py
# Get the directory where this script lives (NOT INTENDED FOR USER EDITING)
SCRIPT_DIR = Path(__file__).parent


# === Consts for script user to alter per-run ===

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
OUTPUT_PPTX_FILENAME = r"sample_slides.pptx"

# CASE-SENSITIVE: Specify the headings prefixes to use when building slide chunks based on headings.
ALLOWED_HEADING_PREFIXES = {
    "Heading",
    "Chapter",
}  # E.g., Heading1, Heading2, Heading3, Chapter1, Chapter2
# You can alter this if your word.docx has custom heading names.
# For chunking based on headings, flat, it doesn't matter what comes after the prefix.
# For chunking based on nested headings, the code assumes there will be a number AT THE END of the name.
# If the number is somewhere else (start, middle), you'll have to adjust the code in (at least) get_style_parts()

# For chunking by NESTED headings, specify their hierarchy below. 1 is the topmost level, 1+N are deeper levels.
HEADING_HIERARCHY = {
    "Chapter": 1,
    "Heading": 2,
    "Scene": 3,
    "Beat": 4,
    # Add more as needed
}

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
# endregion

PRESERVE_COMMENTS: bool = True
PRESERVE_FOOTNOTES: bool = True
PRESERVE_ENDNOTES: bool = False

PRESERVE_ANNOTATIONS: bool = (
    PRESERVE_COMMENTS or PRESERVE_FOOTNOTES or PRESERVE_ENDNOTES
)

COMMENTS_SORT_BY_DATE: bool = True
COMMENTS_KEEP_AUTHOR_AND_DATE: bool = True

FAIL_FAST: bool = False


# region models.py
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

    comments: list[Comment_docx] = field(default_factory=list[Comment_docx])
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

    def add_comment(self, comment: Comment_docx) -> None:
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
# eventual destination: ./src/docxtext2pptx/__main__.py
def main() -> None:
    """Entry point for program flow."""
    setup_console_encoding()
    debug_print("Hello, manuscript parser!")

    user_path = INPUT_DOCX_FILE

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

    if PRESERVE_ANNOTATIONS:
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
    try:
        save_pptx(output_prs)
    except Exception as e:
        print(f"Save failed with error: {e}")
        sys.exit(1)


# endregion


# region create_slides.py
# eventual destination: ./src/docxtext2pptx/create_slides.py
def copy_run_formatting(source_run: Run_docx, target_run: Run_pptx) -> None:
    """Mutates a pptx _Run object to apply text and formatting from a docx Run object."""
    sfont = source_run.font
    tfont = target_run.font

    target_run.text = source_run.text

    # Bold/Italics: Only overwrite when explicitly set on the source (avoid clobbering inheritance)
    if sfont.bold is not None:
        tfont.bold = sfont.bold
    if sfont.italic is not None:
        tfont.italic = sfont.italic

    # Underline: collapse any explicit value (True/False/WD_UNDERLINE.*) to bool
    if sfont.underline is not None:
        tfont.underline = bool(sfont.underline)

    # Color: copy only if source has an explicit RGB
    src_rgb = getattr(getattr(sfont, "color", None), "rgb", None)
    if src_rgb is not None:
        tfont.color.rgb = RGBColor(*src_rgb)


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

        # For each paragraph in this chunk, handle adding it
        for i, paragraph in enumerate(chunk.paragraphs):

            # Creating a new slide and a text frame leaves an empty paragraph in place, even when clearing it.
            # So if we're at the start of our list, use that existing empty paragraph.
            if (
                i == 0
                and len(text_frame.paragraphs) > 0
                and not text_frame.paragraphs[0].text
            ):
                # Use the existing first/0th paragraph
                pptx_paragraph = text_frame.paragraphs[0]
            else:
                pptx_paragraph = text_frame.add_paragraph()

            # Process the docx's paragraph contents, including both runs & hyperlinks
            copy_chunk_paragraph_inner_contents(paragraph, pptx_paragraph)

        if PRESERVE_ANNOTATIONS:
            notes_text_frame: TextFrame = new_slide.notes_slide.notes_text_frame  # type: ignore # TODO, POLISH: is there a reasonable way to fix this type hint?
            annotate_slide(chunk, notes_text_frame)


def copy_chunk_paragraph_inner_contents(
    paragraph: Paragraph_docx, pptx_paragraph: Paragraph_pptx
) -> None:
    """Iterate through a paragraph's runs and hyperlinks, in document order, and process each."""
    items_processed = False

    for item in paragraph.iter_inner_content():
        items_processed = True
        if isinstance(item, Run_docx):

            # If this Run has a field code for instrText and it begins with HYPERLINK, this is an old-style
            # word hyperlink, which we cannot handle the same way as normal docx hyperlinks. But we try to detect
            # when it happens and report it to the user.
            if "instrText" in item.element.xml and "HYPERLINK" in item.element.xml:
                detect_field_code_hyperlinks(item)

            process_run(item, pptx_paragraph)

        elif hasattr(item, "url"):
            # Process all runs within the hyperlink
            for run in item.runs:
                process_run(run, pptx_paragraph, item.url)
        else:
            debug_print(f"Unknown content type in paragraph: {type(item)}")

    # Fallback: if no content was processed but paragraph has text
    if not items_processed and paragraph.text:
        debug_print(
            f"Fallback: paragraph has text but no runs/hyperlinks: {paragraph.text[:50]}"
        )
        pptx_run = pptx_paragraph.add_run()
        pptx_run.text = paragraph.text


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
                debug_print(
                    f"WARNING: We found a field code hyperlink, but we don't have a way to attach it to any text: {instr.text}"
                )

    except (AttributeError, ET.ParseError) as e:
        debug_print(
            f"WARNING: Could not parse run XML for field codes: {e} while seeking instrText"
        )


def process_run(
    run: Run_docx, pptx_paragraph: Paragraph_pptx, hyperlink: str | None = None
) -> Run_pptx:
    """Copy a run from the docx parent to the pptx paragraph, including copying its formatting."""
    # Handle formatting

    pptx_run = pptx_paragraph.add_run()
    copy_run_formatting(run, pptx_run)

    if hyperlink:
        pptx_run_url = pptx_run.hyperlink
        pptx_run_url.address = hyperlink

    return pptx_run


# endregion

# region GET annotation helpers
# eventual destination: ./src/docxtext2pptx/annotations/annotate_chunks.py
# endregion


# region annotate_chunks.py
def process_chunk_annotations(
    chunks: list[Chunk_docx], doc: document.Document
) -> list[Chunk_docx]:
    """For a list of Chunk_docx objects, populate the annotation dicts for each one."""

    # Gather all the doc annotations - use empty dict if feature disabled
    all_comments = get_all_docx_comments(doc) if PRESERVE_COMMENTS else {}
    all_footnotes = get_all_docx_footnotes(doc) if PRESERVE_FOOTNOTES else {}
    all_endnotes = get_all_docx_endnotes(doc) if PRESERVE_ENDNOTES else {}

    for chunk in chunks:
        for paragraph in chunk.paragraphs:
            for item in paragraph.iter_inner_content():

                if isinstance(item, Run_docx):
                    process_run_annotations(
                        chunk,
                        item,
                        all_comments=all_comments,
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
                            run,
                            all_comments=all_comments,
                            all_footnotes=all_footnotes,
                            all_endnotes=all_endnotes,
                        )
                else:
                    debug_print(f"Unknown content type in paragraph: {type(item)}")

    return chunks


def process_run_annotations(
    chunk: Chunk_docx,
    run: Run_docx,
    all_comments: dict[str, Comment_docx],
    all_footnotes: dict[str, Footnote_docx],
    all_endnotes: dict[str, Endnote_docx],
) -> None:
    """Get the annotations from a run object and adding them into its (grand)parent chunk object."""
    try:
        # Get XML from the run using public API
        run_xml = run.element.xml  # type: ignore

        # Parse it safely with ElementTree
        root = ET.fromstring(run_xml)  # type: ignore

        # Define namespace
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # Find comment references
        if PRESERVE_COMMENTS:
            comment_refs = root.findall(".//w:commentReference", ns)
            for ref in comment_refs:
                comment_id = ref.get(f'{{{ns["w"]}}}id')
                if comment_id and comment_id in all_comments:
                    comment_object = all_comments[comment_id]
                    chunk.add_comment(comment_object)

        # Find footnote references
        if PRESERVE_FOOTNOTES:
            footnote_refs = root.findall(".//w:footnoteReference", ns)
            for ref in footnote_refs:
                footnote_id = ref.get(f'{{{ns["w"]}}}id')
                if footnote_id and footnote_id in all_footnotes:
                    footnote_obj = all_footnotes[footnote_id]
                    chunk.add_footnote(footnote_obj)

        # Find endnote references - same pattern
        # if PRESERVE_ENDNOTES:
        #     endnote_refs = root.findall('.//w:endnoteReference', ns)
        #     for ref in endnote_refs:
        #         endnote_id = ref.get(f'{{{ns["w"]}}}id')
        #         if endnote_id and endnote_id in all_endnotes:
        #             endnote_obj = all_endnotes[endnote_id]
        #             chunk.add_endnote(endnote_obj)

    except (AttributeError, ET.ParseError) as e:
        if FAIL_FAST:
            raise
        debug_print(f"WARNING: Could not parse run XML for references: {e}")


# endregion


# eventual destination: ./src/docxtext2pptx/annotations/get_docx_annotations.py
# region Get all comments/footnotes/endnotes in this doc
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
    try:
        # Inspect the docx package as a zip
        zip_package = doc.part.package
        if zip_package is None:
            debug_print("WARNING: Could not access docx package")
            return {}

        footnote_parts: list[Part] = []
        for part in zip_package.parts:
            if "footnotes.xml" in str(part.partname):
                footnote_parts.append(part)

        # If the list is empty, there weren't footnotes.
        if not footnote_parts:
            return {}

        # Parse the footnotes XML

        # We think this will always be a list of one item, so assign that item to a variable.
        footnote_blob = footnote_parts[0].blob

        # If footnote_blob is in bytes, convert it to a string
        if isinstance(footnote_blob, bytes) or hasattr(footnote_blob, "decode"):
            xml_string = bytes(footnote_blob).decode("utf-8")
        else:
            # If it wasn't bytes, we assume it's a string already
            xml_string = footnote_blob

        # Create an ElementTree object by deserializing the footnotes.xml contents into a Python object
        root: ET.Element = ET.fromstring(xml_string)

        return extract_footnotes_from_xml(root)

    except Exception as e:
        if FAIL_FAST:
            raise
        else:
            debug_print(f"Warning: Could not extract footnotes: {e}")
            return {}


def extract_footnotes_from_xml(root: ET.Element) -> dict[str, Footnote_docx]:
    """Extract all footnotes from the doc's footnotes part xml, returning a dict of {id: text}"""

    # We need to define the namespace first
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    footnotes_dict: dict[str, Footnote_docx] = {}

    for footnote in root:
        # Get the footnote ID
        footnote_id = footnote.get(f'{{{ns["w"]}}}id')
        footnote_type = footnote.get(f'{{{ns["w"]}}}type')

        # Skip if we can't get a valid ID
        if footnote_id is None:
            debug_print("WARNING: Found footnote without ID, skipping")
            continue

        # Skip separator footnotes (they're just formatting)
        if footnote_type in ["separator", "continuationSeparator"]:
            continue

        full_text = "".join(footnote.itertext())
        all_hyperlinks = extract_hyperlinks_from_footnote(footnote)

        footnote_obj = Footnote_docx(
            footnote_id=footnote_id, text_body=full_text, hyperlinks=all_hyperlinks
        )

        footnotes_dict[footnote_id] = footnote_obj

    return footnotes_dict


def extract_hyperlinks_from_footnote(element: ET.Element) -> list[str]:
    """Extract all hyperlinks from a footnote element."""
    hyperlinks: list[str] = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    element.iter

    for hyperlink in element.findall(".//w:hyperlink", ns):
        # Get the link text
        link_text = "".join([t.text or "" for t in hyperlink.findall(".//w:t", ns)])
        if link_text.strip():
            hyperlinks.append(link_text.strip())

    return hyperlinks


def get_all_docx_endnotes(doc: document.Document) -> dict[str, Endnote_docx]:
    raise NotImplementedError


def extract_endnotes_from_xml(root: ET.Element) -> dict[str, Endnote_docx]:
    raise NotImplementedError


# endregion


# region annotate_slides.py
# eventual destination: ./src/docxtext2pptx/annotations/annotate_slides.py


def annotate_slide(chunk: Chunk_docx, notes_text_frame: TextFrame) -> None:
    """
    Pull a chunk's preserved annotations and copy them into the slide's speaker notes text frame.

    NOTE: We DO NOT PRESERVE any anchoring to the slide's body for annotations. That means we don't preserve
    comments' selected text ranges, nor do we preserve footnote or endnote numbering.
    """
    if chunk.comments:
        add_comments_to_notes(chunk.comments, notes_text_frame)

    if chunk.footnotes:
        add_footnotes_to_notes(chunk.footnotes, notes_text_frame)

    if chunk.endnotes:
        add_endnotes_to_notes(chunk.endnotes, notes_text_frame)


# region Add annotations to pptx speaker notes text frame
def add_comments_to_notes(
    comments_list: list[Comment_docx], notes_text_frame: TextFrame
) -> None:
    """Copy logic for appending the comments portion of the speaker notes."""

    if comments_list:
        if COMMENTS_SORT_BY_DATE:
            # Sort comments by date (newest first, or change reverse=False for oldest first)
            sorted_comments = sorted(
                comments_list,
                key=lambda c: getattr(c, "timestamp", None) or datetime.min,
                reverse=False,
            )
        else:
            sorted_comments = comments_list

        comment_para = notes_text_frame.add_paragraph()
        comment_run = comment_para.add_run()
        comment_run.text = "COMMENTS FROM SOURCE DOCUMENT:\n" + "=" * 40

        for i, comment in enumerate(sorted_comments, 1):
            # Check if comment has paragraphs attribute
            if hasattr(comment, "paragraphs"):
                # Get paragraphs safely, default to empty list if not present
                this_comment_paragraphs = getattr(comment, "paragraphs", [])
                for para in this_comment_paragraphs:
                    if hasattr(para, "text") and para.text.rstrip():
                        notes_para = notes_text_frame.add_paragraph()
                        comment_header = notes_para.add_run()

                        if COMMENTS_KEEP_AUTHOR_AND_DATE:
                            author = getattr(comment, "author", "Unknown Author")
                            timestamp = getattr(comment, "timestamp", None)

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
                        copy_chunk_paragraph_inner_contents(para, notes_para)


# TODO, ANNOTATIONS: complete for footnotes and endnotes
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
            footnote_run.text = (
                f"\n{footnote_obj.footnote_id}. {footnote_obj.text_body}\n"
            )


def add_endnotes_to_notes(
    endnotes_list: list[Endnote_docx], notes_text_frame: TextFrame
) -> None:
    """Copy logic for appending the endnote portion of the speaker notes."""
    if endnotes_list:
        pass


# endregion


# region create_chunks.py
# eventual destination: ./src/docxtext2pptx/create_chunks.py
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

    # Collect the possible heading-like style_ids in THIS document
    doc_headings = find_numbered_headings(doc)

    if not doc_headings:
        print("No valid numbered headings found for chunk by headings nested.")
        print("Falling back to flat heading chunking...")
        return chunk_by_heading_flat(doc)

    # Find which paragraphs are headings; return the index number and the style_id
    heading_paras = find_heading_indices(doc, doc_headings)

    # Start building the chunks
    all_chunks: list[Chunk_docx] = []
    current_chunk: Chunk_docx = Chunk_docx()

    # Initialize current_heading_style_id  - handle case where no headings exist
    if heading_paras:
        # Set to the style_id of the first-found heading paragraph in the doc
        current_heading_style_id = sorted(heading_paras)[0][1]
    else:
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
            if style_id in doc_headings:
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
            if style_id in doc_headings:
                current_heading_style_id = style_id
            continue

        # Handle headings
        if style_id in doc_headings:
            # Check if this heading is at same level or higher (less deep) than current
            if get_heading_depth(style_id) <= get_heading_depth(
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


## === chunk_by_heading_nested() helpers ===


def get_style_parts(style_id: str) -> tuple[str, int] | None:
    """Split the style_id into prefix & number using RegEx"""
    match = re.match(r"([A-Za-z]+)(\d+)$", style_id)
    if match:
        style_prefix, style_num_str = match.groups()
        style_num = int(style_num_str)
        return (style_prefix, style_num)
    return None


def find_doc_prefixed_headings(doc: document.Document) -> set[str]:
    """Generate the set of style_ids used by paragraphs in this document that match the approved prefixes list."""
    styles_found: set[str] = set()

    for para in doc.paragraphs:

        # If this paragraph doesn't even have a style_id then we don't care; skip it
        if not (para.style and para.style.style_id):
            continue
        style_id = para.style.style_id

        # # Split out the prefix and see if it is in the allowed headings prefixes
        if style_id.startswith(tuple(ALLOWED_HEADING_PREFIXES)):
            styles_found.add(style_id)

    debug_print(sorted(styles_found))
    return styles_found


def find_numbered_headings(doc: document.Document) -> list[str] | None:
    """Find headings used in this document that end in a number."""
    all_headings: set[str] = find_doc_prefixed_headings(doc)
    numbered_headings: list[str] = []
    for heading in all_headings:
        style_parts = get_style_parts(heading)
        if style_parts:
            numbered_headings.append(heading)
    if numbered_headings:
        return numbered_headings
    else:
        return None


def find_heading_indices(
    doc: document.Document, headings: list[str]
) -> list[tuple[int, str]]:
    """Find all paragraphs in this document that are headings and store their (index, style_id) in a set."""
    heading_paragraphs: set[tuple[int, str]] = set()
    for i, para in enumerate(doc.paragraphs):
        if para.style:
            style_id = para.style.style_id
        else:
            continue
        if style_id in headings:
            heading_paragraphs.add((i, style_id))
    return sorted(heading_paragraphs)


def get_heading_depth(style_id: str) -> int | float:
    """Compute a heading depth based on the defined hierarchy and style_id number."""
    style_parts = get_style_parts(style_id)

    if not style_parts:
        return float("inf")

    style_prefix, style_num = style_parts

    # Look up this heading's prefix in the hierarchy list. If it's not there, default it to 99 so it's low-pri.
    # (Enough checks prior to this call should prevent that case, but let's not just assume I got the logic right...)
    hierarchy_depth = HEADING_HIERARCHY.get(style_prefix, 99)

    # Multiply the hierarchy depth by 1000, then add the style's depth.
    # E.g., if Chapter = 1, Heading = 2, then:
    # Chapter2 = 1002
    # Heading1 = 2001
    # Heading12 = 2012
    # Safe up to heading numbers < 1000 (current max expected: ~15)
    comparison_depth = (hierarchy_depth * 1000) + style_num

    return comparison_depth


## ===

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

    # Collect the possible heading-like style_ids in THIS document
    doc_headings = find_doc_prefixed_headings(doc)

    if not doc_headings:
        print(
            f"Warning: No headings found matching prefixes {ALLOWED_HEADING_PREFIXES}"
        )
        print("Falling back to paragraph chunking...")
        return chunk_by_paragraph(doc)

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
        if style_id in doc_headings:
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


# TODO: cut?
def fail_or_warn(
    error_msg: str,
    warning_msg: str,
    exception_type: type[Exception] = ValueError,
    fail_fast: bool = False,
) -> None:
    """Handle error conditions with configurable fail-fast behavior."""
    if fail_fast:
        raise exception_type(error_msg)
    else:
        debug_print(f"WARNING: {warning_msg}")


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
# eventual destination: ./src/docxtext2pptx/io.py (?)
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
def save_pptx(prs: presentation.Presentation) -> None:
    """Save the generated slides to disk."""
    try:
        # Construct output path
        if OUTPUT_PPTX_FOLDER:
            folder = Path(OUTPUT_PPTX_FOLDER)
            folder.mkdir(parents=True, exist_ok=True)
            output_filepath = folder / OUTPUT_PPTX_FILENAME
        else:
            output_filepath = Path(OUTPUT_PPTX_FILENAME)
        prs.save(str(output_filepath))
        print(f"Successfully saved to {output_filepath}")
    except PermissionError:
        raise PermissionError("Save failed: File may be open in another program")
    except Exception as e:
        raise RuntimeError(f"Save failed with error: {e}")


# endregion

# region call main
if __name__ == "__main__":
    main()
# endregion
