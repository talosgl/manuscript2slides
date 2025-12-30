"""Extract annotations from Word documents."""

import logging
import re
import xml.etree.ElementTree as ET
from typing import TypeVar

from docx import document
from docx.comments import Comment as Comment_docx
from docx.text.paragraph import Paragraph as Paragraph_docx
from docx.text.run import Run as Run_docx

from manuscript2slides.internals.define_config import UserConfig
from manuscript2slides.models import (
    Chunk_docx,
    Comment_docx_custom,
    Endnote_docx,
    Footnote_docx,
)
from manuscript2slides.processing import docx_xml

log = logging.getLogger("manuscript2slides")
NOTE_TYPE = TypeVar("NOTE_TYPE", Footnote_docx, Endnote_docx)


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

        # Strip leading footnote/endnote number (e.g. "1. text" -> "text" or "1 text" -> "text")
        # The period after the number is optional as some Word versions don't include it
        note_full_text = re.sub(r"^\d+\.?\s*", "", note_full_text)

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


# region process_chunk_annotations
def process_chunk_annotations(
    chunks: list[Chunk_docx], doc: document.Document, cfg: UserConfig
) -> list[Chunk_docx]:
    """For a list of Chunk_docx objects, populate the annotation dicts for each one."""

    # Gather all the doc annotations
    all_raw_comments = get_all_docx_comments(doc)
    all_footnotes = get_all_docx_footnotes(doc, cfg)
    all_endnotes = get_all_docx_endnotes(doc, cfg)

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
                    log.warning(f"Unknown content type in paragraph: {type(item)}")

    return chunks


# endregion


# region process_run_annotations
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
        run_xml = run.element.xml

        # Parse it safely with ElementTree
        root = ET.fromstring(run_xml)

        # Define namespace
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # get the reference text to be used by comments, footnotes, or endnotes
        ref_text = get_ref_text(run, paragraph)

        # Find comment references
        comment_refs = root.findall(".//w:commentReference", ns)
        for ref in comment_refs:
            comment_id = ref.get(f"{{{ns['w']}}}id")
            if comment_id and comment_id in all_raw_comments:
                comment_object = all_raw_comments[comment_id]

                custom_comment_obj = Comment_docx_custom(
                    comment_obj=comment_object, reference_text=ref_text
                )
                chunk.add_comment(custom_comment_obj)

        # Find footnote references
        footnote_refs = root.findall(".//w:footnoteReference", ns)
        for ref in footnote_refs:
            footnote_id = ref.get(f"{{{ns['w']}}}id")
            if footnote_id and footnote_id in all_footnotes:
                footnote_obj = all_footnotes[footnote_id]
                footnote_obj.reference_text = ref_text
                chunk.add_footnote(footnote_obj)

        # Find endnote references - same pattern
        endnote_refs = root.findall(".//w:endnoteReference", ns)
        for ref in endnote_refs:
            endnote_id = ref.get(f"{{{ns['w']}}}id")
            if endnote_id and endnote_id in all_endnotes:
                endnote_obj = all_endnotes[endnote_id]
                endnote_obj.reference_text = ref_text
                chunk.add_endnote(endnote_obj)

    except (AttributeError, ET.ParseError) as e:
        log.warning(f"Could not parse run XML for references: {e}")


# endregion


# region get_ref_text
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


# endregion


# region get_all_docx_comments
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


# endregion


# region get_all_docx_footnotes
def get_all_docx_footnotes(
    doc: document.Document, cfg: UserConfig
) -> dict[str, Footnote_docx]:
    """
    Extract all footnotes from a docx document.
    Returns {id: {footnote_id: str, text_body: str, hyperlinks: list of str} }.
    """

    if not cfg.display_footnotes:
        return {}

    try:
        footnotes_parts = docx_xml.find_xml_parts(doc, "footnotes.xml")

        if not footnotes_parts:
            return {}

        # We think this will always be a list of one item, so assign that item to a variable.
        root = docx_xml.parse_xml_blob(footnotes_parts[0].blob)
        return extract_notes_from_xml(root, Footnote_docx)

    except Exception as e:
        log.warning(f"Could not extract footnotes: {e}")
        return {}


# endregion


# region get_all_docx_endnotes
def get_all_docx_endnotes(
    doc: document.Document, cfg: UserConfig
) -> dict[str, Endnote_docx]:
    """
    Extract all endnotes from a docx document.
    Returns {id: {footnote_id: str, text_body: str, hyperlinks: list of str} }.
    """
    if not cfg.display_endnotes:
        return {}

    try:
        endnotes_parts = docx_xml.find_xml_parts(doc, "endnotes.xml")

        if not endnotes_parts:
            return {}

        root = docx_xml.parse_xml_blob(endnotes_parts[0].blob)
        return extract_notes_from_xml(root, Endnote_docx)

    except Exception as e:
        log.warning(f"Could not extract endnotes: {e}")
        return {}


# endregion
