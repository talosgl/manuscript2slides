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
from typing import Union
from dataclasses import dataclass, field

# Third-party libraries
import docx
import pptx
from docx import document
from docx.comments import Comment as Comment_docx

# from docx.table import Table as Table_docx
# from docx.text.hyperlink import Hyperlink as Hyperlink_docx
from docx.text.paragraph import Paragraph as Paragraph_docx
from docx.text.run import Run as Run_docx
from pptx import presentation
from pptx.dml.color import RGBColor
from pptx.slide import Slide, SlideLayout
from pptx.text.text import TextFrame, _Paragraph as Paragraph_pptx, _Run as Run_pptx  # type: ignore
from pptx.shapes.placeholder import SlidePlaceholder

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
    
    -   REVERSE FLOW LIMITATIONS
        -   The reverse flow (pptx2docx) is significantly less robust for annotations. There is no support for separating out
            speaker notes into types of annotations. "We simply take all speaker notes and attach them to the final paragraph 
            of each slide's content as a comment. Original source docx comment metadata is not preserved (apart from
            plain text author/timestamps if kept during an original docx2pptx conversion).

"""
# endregion

# region CONSTANTS
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
PRESERVE_FOOTNOTES: bool = False
PRESERVE_ENDNOTES: bool = False

PRESERVE_ANNOTATIONS: bool = (
    PRESERVE_COMMENTS or PRESERVE_FOOTNOTES or PRESERVE_ENDNOTES
)

COMMENTS_SORT_BY_DATE: bool = True
COMMENTS_KEEP_AUTHOR_AND_DATE: bool = True

# region Classes


@dataclass
class Footnote_docx:
    footnote_id: str
    paragraphs: list[Paragraph_docx] = field(default_factory=list[Paragraph_docx])


@dataclass
class Endnote_docx:
    endnote_id: str
    paragraphs: list[Paragraph_docx] = field(default_factory=list[Paragraph_docx])


# Type aliases for clarity
# ParagraphType = Union[Paragraph_docx, Paragraph_pptx] # might allow us to reuse functions for either-direction import/export (I hope)
# InnerContentType_docx = Paragraph_docx | Table_docx # might allow us to support tables and paragraphs
AnnotationType = Union[Comment_docx, Footnote_docx, Endnote_docx]


@dataclass
class Chunk_docx:
    """Class for Chunk objects made from docx paragraphs and their associated annotations."""

    # Use "default_factory" to ensure every chunk gets its own list.
    paragraphs: list[Paragraph_docx] = field(default_factory=list[Paragraph_docx])

    # TODO, TABLE SUPPORT: if we support tables later
    # inner_contents: list[InnerContentType_docx] = field(default_factory=list)

    annotations: dict[str, list[AnnotationType]] = field(  # type: ignore
        default_factory=lambda: {
            "comments": [],
            "footnotes": [],
            "endnotes": [],
        }  # Calls dict() constructor with our key-value pairs
    )

    @classmethod
    def create_with_paragraph(cls, paragraph: Paragraph_docx) -> "Chunk_docx":
        """Create a new instance of a Chunk_docx object but also populate the paragraphs list with an initial element."""
        return cls(paragraphs=[paragraph])

    def add_annotation(self, annotation_type: str, annotation: AnnotationType) -> None:
        """Add an annotation to the appropriate list for this instance of a Chunk object."""
        if annotation_type not in self.annotations:
            self.annotations[annotation_type] = []
        self.annotations[annotation_type].append(annotation)

    def add_paragraph(self, new_paragraph: Paragraph_docx) -> None:
        """Add a paragraph to this Chunk object's paragraphs list."""
        self.paragraphs.append(new_paragraph)

    def add_paragraphs(self, new_paragraphs: list[Paragraph_docx]) -> None:
        """Add a list of paragraphs to this Chunk object's paragraphs list."""
        self.paragraphs.extend(new_paragraphs)  # Add multiple at once


# TODO, REVERSE FLOW: We probably need a separate Chunk_pptx class
# I think it would be:
# list of body [Paragraph_pptx]
# list of speaker notes [Paragraph_pptx]
# make sure to use a factory for these lists because lists are mutable!


# endregion


# region Main program flow
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


# region Pipeline Functions
# TODO, BASIC VALIDATION: Add basic validation for docx contents (copy the validation we did in CSharp)
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


# region Slide creation helpers


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


# endregion


# region process_paragraph_inner_contents (Chunk_docx)
def process_paragraph_inner_contents(
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
<<<<<<< HEAD
            # TODO, POLISH: fix this code to more safely handle unexpected XML structures (if hasattr() etc.)
            # TODO, POLISH: maybe put this into a helper function handle_instrText_hyperlinks() or something?
=======
>>>>>>> 7de168d73cdacc1302b4db453b9adb6363189977
            for child in item._element:  # type: ignore
                if (
                    child.tag.endswith("instrText")  # type: ignore
                    and child.text  # type: ignore
                    and child.text.startswith("HYPERLINK")  # type: ignore
                ):
                    debug_print(
                        f"WARNING: We found a field code hyperlink, but we don't have a way to attach it to any text: {child.text}"  # type: ignore
                    )

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


# endregion


# region process_run (Chunk_docx)
def process_run(
    run: Run_docx, pptx_paragraph: Paragraph_pptx, hyperlink: str | None = None
) -> Run_pptx:
    """Process a run, including copying its formatting."""
    # Handle formatting

    pptx_run = pptx_paragraph.add_run()
    copy_run_formatting(run, pptx_run)

    if hyperlink:
        pptx_run_url = pptx_run.hyperlink
        pptx_run_url.address = hyperlink

    return pptx_run


# endregion


# region annotation helpers (Chunk_docx)
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


# TODO, ANNOTATIONS: split up into 3 calls: get_docx_comments, get_docx_footnotes, get_docx_endnotes, and use this as orchestrator
def process_chunk_annotations(
    chunks: list[Chunk_docx], doc: document.Document
) -> list[Chunk_docx]:
    """For a list of Chunk_docx objects, populate the annotation dicts for each one."""
    # Gather all the doc annotations
    all_comments = get_all_docx_comments(doc)

    # TODO, ANNOTATIONS: implement get_all_docx_footnotes and get_all_docx_endnotes
    # all_footnotes = []
    # all_endnotes = []

    for chunk in chunks:
        for paragraph in chunk.paragraphs:
            for item in paragraph.iter_inner_content():
                for child in item._element:  # type: ignore[attr-defined]
                    # PROCESS COMMENTS

                    # Check if child has the attributes we need
                    if not (hasattr(child, "tag") and hasattr(child, "get")):  # type: ignore[attr-defined]
                        continue

                    if PRESERVE_COMMENTS and child.tag.endswith("commentReference"):  # type: ignore[attr-defined]
                        # Extract the id attribute
                        new_comment_id = child.get(  # type: ignore[attr-defined]
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id"
                        )
                        # item_comment_ids.append(new_comment_id)
                        if new_comment_id and new_comment_id in all_comments:
                            comment_object = all_comments[new_comment_id]
                            chunk.add_annotation("comments", comment_object)

                # TODO, ANNOTATIONS: complete this for footnotes and endnotes
                if PRESERVE_FOOTNOTES and child.tag.endswith("footnoteReference"):  # type: ignore[attr-defined]
                    pass
                #     new_footnote_id = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                #     item_footnote_ids.append(new_footnote_id)

                if PRESERVE_ENDNOTES and child.tag.endswith("endnoteReference"):  # type: ignore[attr-defined]
                    pass
                #     new_endnote_id = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                #     item_endnote_ids.append(new_endnote_id)

    return chunks


# TODO, ANNOTATIONS: as above-- split up into 3 calls: get_chunk_comments, get_chunk_footnotes, get_chunk_endnotes, and use this as orchestrator
def annotate_chunk(chunk: Chunk_docx, notes_text_frame: TextFrame) -> Chunk_docx:
    """
    Pull this chunk's preserved annotations and format them for a slide's speaker notes.

    NOTE: We DO NOT PRESERVE any anchoring to the slide's body for annotations. That means we don't preserve
    comments' selected text ranges, nor do we preserve footnote or endnote numbering.
    """
    comments_list = chunk.annotations.get("comments", [])

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
                        process_paragraph_inner_contents(para, notes_para)

    # TODO, ANNOTATIONS: complete this for footnotes and endnotes
    footnotes_list = chunk.annotations.get("footnotes", [])
    if footnotes_list:
        footnote_section = "FOOTNOTES FROM SOURCE DOCUMENT:\n" + "=" * 40 + "\n\n"
        # TODO: I don't know about using i here; footnote starting numbers are gonna change page over page
        for i, footnote_item in enumerate(footnotes_list, 1):

            # Handle different possible footnote formats
            if isinstance(footnote_item, tuple) and len(footnote_item) >= 2:
                footnote_id, footnote_text = footnote_item[0], footnote_item[1]  # type: ignore (while in dev)
            else:
                footnote_id = i
                footnote_text = str(footnote_item)
            footnote_section += f"{i}. Footnote {footnote_id}:\n"
            footnote_section += f"   {footnote_text}\n\n"

    endnotes_list = chunk.annotations.get("endnotes", [])
    if endnotes_list:
        pass

    return chunk


# endregion


# region slides_from_chunks
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
            process_paragraph_inner_contents(paragraph, pptx_paragraph)

        if PRESERVE_ANNOTATIONS:
            notes_text_frame: TextFrame = new_slide.notes_slide.notes_text_frame  # type: ignore # TODO, POLISH: is there a reasonable way to fix this type hint?
            annotate_chunk(chunk, notes_text_frame)


# endregion


# region === Chunking Functions ===

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
        current_page_chunk.add_paragraph(
            para
        )

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
