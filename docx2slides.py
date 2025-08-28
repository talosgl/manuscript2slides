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

from __future__ import annotations

import sys
import docx
from docx import document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
import pptx
from pptx import presentation
from pptx.text.text import _Run  # type: ignore
from pptx.dml.color import RGBColor
from pathlib import Path
import re
import platform
import io


# region Overarching TODOs
"""
Features to Add:
v1 (additive features)
- Polish up enough to make repo public
- What does is "done enough for public github repo mean"? "ready when I'm comfortable having strangers use it without asking me questions." 
    - Error messages that tell users what went wrong and how to fix it
    - Code that doesn't crash on common edge cases.

Wishlist Features for Python version:
- Add support for importing .doc (can't use python-docx and will need to learn and use python-docbinary)
- Add support for importing .md and .txt; split by whitespaces or character like \n\n.
- Add support for outputting to other slide formats that aren't pptx
- Add support to break chunks (of any type) at a word count threshold


Features deferred to C# Rewrite:
- Reverse flow: export slide text frame content to docx paragraphs with the annotations kept as comments, inline
    - This could allow a flow where you can iterate back and forth (and remove the need to "update" an existing deck with manuscript updates)
- Retain comments for a block of text and put them into the speaker notes for the slide where that block goes
- Retain footnotes & endnotes
- Build a simple UI + Package this so that any writer can use it without needing to know WTF python is
    - + Consider reworking into C# for this; good practice for me

"""
# endregion

# region CONSTANTS
# Get the directory where this script lives (NOT INTENDED FOR USER EDITING)
SCRIPT_DIR = Path(__file__).parent


# === Consts for script user to alter per-run ===

# The pptx file to use as the template for the slide deck
TEMPLATE_PPTX = SCRIPT_DIR / "resources" / "blank_pptx_landscape_notes_view.pptx"
# You can make your own template with the master slide and master notes page
# to determine how the output will look. You can customize things like font, paragraph style,
# slide size, slide layout...

# Desired slide layout. All slides use the same layout.
SLD_LAYOUT_CUSTOM = 11
# This is an index id. From the master slides template, count from 0, 1, 2...

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

# Which chunking method to use to divide the docx into slides
CHUNK_TYPE = "heading_flat"  # Options: "heading_nested", "heading_flat", "paragraph", "page" -- see create_docx_chunks()

# Toggle on/off whether to print debug_prints() to the console
DEBUG_MODE = True  # TODO, v1 POLISH: set to false before publishing
# endregion

# Specify whether to keep text formatting like italics, bold, color. Skipping can make the script go faster.
KEEP_FORMATTING = True


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
    # chunks = create_docx_chunks(user_docx, "paragraph")
    # chunks = create_docx_chunks(user_docx, "page")
    # chunks = create_docx_chunks(user_docx, "heading_nested")
    chunks = create_docx_chunks(user_docx, CHUNK_TYPE)

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
    except PermissionError:
        print("Save failed: File may be open in another program")
        sys.exit(1)
    except Exception as e:
        print(f"Save failed with error: {e}")
        sys.exit(1)


# endregion


# region Pipeline Functions
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
    doc: document.Document, chunk_type: str = "paragraph"
) -> list[list[Paragraph]]:
    """
    Orchestrator function to create chunks (that will become slides) from the document
    contents, either from paragraph, heading (heading_nested or heading_flat),
    or page. Defaults to paragraph.
    """
    if chunk_type == "heading_flat":
        chunks = chunk_by_heading_flat(doc)
    elif chunk_type == "heading_nested":
        chunks = chunk_by_heading_nested(doc)
    elif chunk_type == "page":
        chunks = chunk_by_page(doc)
    else:
        chunks = chunk_by_paragraph(doc)
    return chunks


def copy_run_formatting(source_run: Run, target_run: _Run) -> None:
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


def slides_from_chunks(
    doc: document.Document,
    prs: presentation.Presentation,
    chunks: list[list[Paragraph]],
) -> None:
    """Generate slide objects, one for each chunk created by earlier pipeline steps."""
    # For every paragraph in the docx that has content, create a slide and populate the content.text with the paragraph's text.
    if SLD_LAYOUT_CUSTOM >= len(prs.slide_layouts):
        raise ValueError(
            f"Slide layout index {SLD_LAYOUT_CUSTOM} doesn't exist. Template has {len(prs.slide_layouts)} layouts (0-{len(prs.slide_layouts)-1})"
        )

    slide_layout = prs.slide_layouts[SLD_LAYOUT_CUSTOM]

    for chunk in chunks:
        # debug_print(f"Creating slide with {len(chunk)} paragraphs, total length: {len(body)} characters")
        # debug_print(f"First 100 chars: {body[:100]}...")
        new_slide = prs.slides.add_slide(slide_layout)
        content = new_slide.placeholders[1]
        text_frame = content.text_frame  # type:ignore
        text_frame.clear()  # type:ignore

        if not KEEP_FORMATTING:
            body = "\n".join(para.text for para in chunk)
            text_frame.text = body
        else:
            for paragraph in chunk:
                pptx_paragraph = text_frame.add_paragraph()  # type:ignore
                for run in paragraph.runs:
                    if run.text:
                        pptx_run = pptx_paragraph.add_run()  # type:ignore
                        copy_run_formatting(run, pptx_run)  # type:ignore

        # add an empty notes area to the slide for annotations
        notes_slide_ptr = new_slide.notes_slide  # type: ignore # noqa


# TODO: The "best" way to globally auto-fit all text in all slides in the PPTX UI is to go to the
# slide master and toggle the auto-fit setting on/off; that'll force all slides to fit the text to the
# textbox frame. So experiment here and see if, after populating all the slides with text, we can then
# alter the master slide and get it to apply the new setting to existing slides. Ideally we'd be able to
# run this BEFORE saving the prs object to .pptx on disk, but maybe if we can't, we can reopen what we just
# saved and save over it?
def resize_text_in_slides(prs: presentation.Presentation) -> None:
    """Auto-fit text to the slides' text-frames."""
    pass
    # PREVIOUS ATTEMPT BELOW
    # TODO, FIX: This appropriately sets the property to auto size the text to fit the shape,
    # but it doesn't actually auto-size it. Presumably inside of the PPTX UI it applies it at render-time.
    # Is there a way to do it programmatically, or not?
    # from pptx.enum.text import MSO_AUTO_SIZE
    # content.text_frame.auto_size, content.text_frame.word_wrap  = (MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, True) # type:ignore
    # content.text_frame._apply_fit("Bookerly", 12, False, False)


# endregion

# region === Chunking Functions ===
# endregion


# region by Paragraph
def chunk_by_paragraph(doc: document.Document) -> list[list[Paragraph]]:
    """
    Creates chunks (which will become slides) based on paragraph, which are blocks of content
    separated by whitespace.
    """
    paragraphs: list[list[Paragraph]] = []
    for para in doc.paragraphs:
        # If this paragraph has no text (whitespace break), skip it
        if para.text == "":
            continue
        para_list = [para]
        paragraphs.append(para_list)
    return paragraphs


# endregion


# region by Page
def chunk_by_page(doc: document.Document) -> list[list[Paragraph]]:
    """Creates chunks based on page breaks"""

    # Start building the chunks
    chunks: list[list[Paragraph]] = []
    current_chunk: list[Paragraph] = []

    for para in doc.paragraphs:
        if para.text == "":
            continue

        # If the current_chunk is empty, append the current para regardless of style & continue to next para.
        if not current_chunk:
            current_chunk.append(para)
            continue

        # Handle page breaks - create new chunk and start fresh
        if para.contains_page_break:
            # Add the current_chunk to chunks list (if it has content)
            if current_chunk:
                chunks.append(current_chunk)
            # Start new chunk with this paragraph
            current_chunk = [para]
            continue

        # If there was no page break, just append this paragraph to the current_chunk
        current_chunk.append(para)

    # Ensure final chunk from loop is added to chunks list
    if current_chunk:
        chunks.append(current_chunk)

    print(f"This document has {len(chunks)} page chunks.")
    return chunks


# endregion

# region by Heading (nested)


def chunk_by_heading_nested(doc: document.Document) -> list[list[Paragraph]]:
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
    chunks: list[list[Paragraph]] = []
    current_chunk: list[Paragraph] = []

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

        # Make Pylance happy (it gets mad if we direct-check para.style.style_id later)
        style_id = para.style.style_id if para.style else "Normal"

        debug_print(f"Paragraph begins: {para.text[:30]}... and is index: {i}")

        # If the current_chunk is empty, append the current para regardless of style & continue to next para.
        if not current_chunk:
            current_chunk.append(para)
            if style_id in doc_headings:
                current_heading_style_id = style_id
            continue

        # Handle page breaks - create new chunk and start fresh
        if para.contains_page_break:
            # Add the current chunk to chunks list (if it has content)
            if current_chunk:
                chunks.append(current_chunk)
            # Start new chunk with this paragraph
            current_chunk = [para]
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
                    chunks.append(current_chunk)
                current_chunk = [para]
                current_heading_style_id = style_id
            else:
                # This heading is deeper, add to current chunk
                current_chunk.append(para)
                current_heading_style_id = style_id
        else:
            # Normal paragraph - add to current chunk
            current_chunk.append(para)

    if current_chunk:
        chunks.append(current_chunk)

    print(f"This document has {len(chunks)} nested heading chunks.")
    return chunks


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
def chunk_by_heading_flat(doc: document.Document) -> list[list[Paragraph]]:
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
    chunks: list[list[Paragraph]] = []
    current_chunk: list[Paragraph] = []

    for para in doc.paragraphs:
        # Skip empty paragraphs
        if para.text == "":
            continue

        # Make Pylance happy (it gets mad if we direct-check para.style.style_id later)
        style_id = para.style.style_id if para.style else "Normal"

        # debug_print(f"Paragraph begins: {para.text[:30]}... and is index: {i}")

        # If the current_chunk is empty, append the current para regardless of style & continue to next para.
        if not current_chunk:
            current_chunk.append(para)
            continue

        # Handle page breaks - always start a new chunk
        if para.contains_page_break:
            # Add the current chunk to chunks list (if it has content)
            if current_chunk:
                chunks.append(current_chunk)
            # Start new chunk with this paragraph
            current_chunk = [para]
            continue

        # If this paragraph is a heading, start a new chunk
        if style_id in doc_headings:
            # If we already have content in current_chunk, save it and start fresh
            if current_chunk:
                chunks.append(current_chunk)
            # Start new chunk with this heading
            current_chunk = [para]
        else:
            # This is a normal paragraph - add it to current chunk
            current_chunk.append(para)

    if current_chunk:
        chunks.append(current_chunk)

    print(f"This document has {len(chunks)} flat heading chunks.")
    return chunks


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

    return path


def validate_pptx_path(user_path: str | Path) -> Path:
    """Validates the pptx template filepath exists and is actually a pptx file."""
    path = validate_path(user_path)
    # Verify it's the right extension:
    if path.suffix.lower() != ".pptx":
        raise ValueError(f"Expected a .pptx file, but got: {path.suffix}")
    return path


def create_empty_slide_deck() -> presentation.Presentation:
    """Load the PowerPoint template and create a new presentation object."""
    template_path = validate_pptx_path(Path(TEMPLATE_PPTX))

    # === testing
    # slide_layout = prs.slide_layouts[SLD_LAYOUT_CUSTOM]
    # slide = prs.slides.add_slide(slide_layout)
    # content = slide.placeholders[1]
    # content.text = "Test Slide!" # type:ignore

    prs = pptx.Presentation(str(template_path))
    return prs


def save_pptx(prs: presentation.Presentation) -> None:
    """Save the generated slides to disk."""
    # Construct output path
    if OUTPUT_PPTX_FOLDER:
        folder = Path(OUTPUT_PPTX_FOLDER)
        folder.mkdir(parents=True, exist_ok=True)
        output_filepath = folder / OUTPUT_PPTX_FILENAME
    else:
        output_filepath = Path(OUTPUT_PPTX_FILENAME)

    prs.save(str(output_filepath))
    print(f"Successfully saved to {output_filepath}")


# endregion

# region call main
if __name__ == "__main__":
    main()
# endregion
