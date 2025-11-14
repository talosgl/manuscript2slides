"""Restore annotations from slide metadata."""

import json
import logging

from manuscript2slides.internals.constants import (
    METADATA_MARKER_FOOTER,
    METADATA_MARKER_HEADER,
    NOTES_MARKER_FOOTER,
    NOTES_MARKER_HEADER,
)
from manuscript2slides.models import SlideNotes

log = logging.getLogger("manuscript2slides")


# region split_speaker_notes
def split_speaker_notes(speaker_notes_text: str) -> SlideNotes:
    """
    Extract user notes, JSON metadata, and discard copied annotations.
    Returns a SlideNotes object with parsed sections (or an empty object if parsing failed).
    """

    # Find all marker positions
    json_start = speaker_notes_text.find(METADATA_MARKER_HEADER)
    json_end = speaker_notes_text.find(METADATA_MARKER_FOOTER)

    notes_start = speaker_notes_text.find(NOTES_MARKER_HEADER)
    notes_end = speaker_notes_text.find(NOTES_MARKER_FOOTER)

    # Extract JSON if present
    json_content = None
    if json_start != -1 and json_end != -1:
        json_text = speaker_notes_text[
            json_start + len(METADATA_MARKER_HEADER) : json_end
        ]
        # Clean up both ends - remove separator lines
        json_text = json_text.strip().lstrip("=").rstrip("=").strip()
        try:
            json_content = json.loads(json_text)
        except json.JSONDecodeError:
            json_content = None
            log.warning(
                "We found what looked like a JSON section but then failed to load it."
            )

    # Build list of ranges to remove (inclusive of markers)
    ranges_to_remove = []

    # If we found a range of JSON start/end, mark that for removal
    if json_start != -1 and json_end != -1:
        ranges_to_remove.append((json_start, json_end + len(METADATA_MARKER_FOOTER)))

    # If we found a range of notes start/end, mark that for removal
    if notes_start != -1 and notes_end != -1:
        ranges_to_remove.append((notes_start, notes_end + len(NOTES_MARKER_FOOTER)))

    # Sort ranges by start position (for consistent removal)
    ranges_to_remove.sort()

    # Extract user notes by removing the marked sections
    user_notes = remove_ranges_from_text(speaker_notes_text, ranges_to_remove)

    # Create and populate the slide_notes object
    slide_notes = SlideNotes()
    slide_notes.user_notes = user_notes.strip()

    if json_content:
        slide_notes = extract_slide_metadata(json_content, slide_notes)

    return slide_notes


# endregion


# region remove_ranges_from_text
def remove_ranges_from_text(text: str, ranges: list) -> str:
    """Remove multiple ranges from text, working backwards to preserve positions."""

    # First merge any overlapping ranges
    ranges_merged = merge_overlapping_ranges(ranges)

    # Sort ranges by start position, then work backwards
    ranges_sorted = sorted(ranges_merged, reverse=True)  # Start from end

    result = text
    for start, end in ranges_sorted:
        result = result[:start] + result[end:]

    return result


# endregion


# region merge_overlapping_ranges
def merge_overlapping_ranges(ranges: list) -> list:
    """If any (int, int) index ranges overlap, merge them."""
    if not ranges:
        return []

    # Sort by start position
    sorted_ranges = sorted(ranges)
    merged = [sorted_ranges[0]]

    for current in sorted_ranges[1:]:
        last = merged[-1]

        # Check if current overlaps with last merged range
        if current[0] <= last[1]:  # start of current <= end of last
            # Merge: extend the end position if needed
            merged[-1] = (last[0], max(last[1], current[1]))
        else:
            # No overlap, add as new range
            merged.append(current)

    return merged


# endregion


# region extract metadata from slide notes
def extract_slide_metadata(json_metadata: dict, slide_notes: SlideNotes) -> SlideNotes:
    """Extract metadata with safe defaults. Add to/mutate slide_notes object and return that object."""

    # Defensively validate the json is a dict.
    if not isinstance(json_metadata, dict):
        # Even though we specified a dict type hint in the func signature, those are for static analysis;
        # they tell Pylance what we expect the type to be, but they don't enforce anything at runtime.
        # So someone could pass a string or list to this function at runtime and our code would still try
        # to use it without this check.

        log.debug(
            f"JSON metadata from this slide should be a dict, but is a {type(json_metadata)}, so we can't use it."
        )
        # return unmutated
        return slide_notes

    # Validate each of the items inside the JSON dict are the expected type (list)
    slide_comments = json_metadata.get("comments", [])
    if not isinstance(slide_comments, list):
        log.debug(
            f"Comments from the slide notes JSON should be a list, but is a {type(slide_comments)}, so we can't use it."
        )
        slide_comments = []

    slide_footnotes = json_metadata.get("footnotes", [])
    if not isinstance(slide_footnotes, list):
        log.debug(
            f"Footnotes from the slide notes JSON should be a list, but is a {type(slide_footnotes)}, so we can't use it."
        )
        slide_footnotes = []

    slide_endnotes = json_metadata.get("endnotes", [])
    if not isinstance(slide_endnotes, list):
        log.debug(
            f"Endnotes from the slide notes JSON should be a list, but is a {type(slide_endnotes)}, so we can't use it."
        )
        slide_endnotes = []

    slide_headings = json_metadata.get("headings", [])
    if not isinstance(slide_headings, list):
        log.debug(
            f"Headings from the slide notes JSON should be a list, but is a {type(slide_headings)}, so we can't use it."
        )
        slide_headings = []

    slide_exp_formatting = json_metadata.get("experimental_formatting", [])
    if not isinstance(slide_exp_formatting, list):
        log.debug(
            f"Experimental_formatting from the slide notes JSON should be a list, but is a {type(slide_exp_formatting)}, so we can't use it."
        )
        slide_exp_formatting = []

    # Populate the slide_notes object with each validated item
    slide_notes.metadata = json_metadata
    slide_notes.comments = slide_comments
    slide_notes.footnotes = slide_footnotes
    slide_notes.endnotes = slide_endnotes
    slide_notes.headings = slide_headings
    slide_notes.experimental_formatting = slide_exp_formatting

    return slide_notes


# endregion


# region safely_extract_comment_data
def safely_extract_comment_data(comment: dict) -> dict | None:
    """
    Extract comment data with validation. Returns None if invalid.
    Returns dict with extracted fields if valid.
    """

    if not isinstance(comment, dict):
        log.debug(f"Comment should be a dict, but is: {type(comment)}.")
        return None

    # Validate the individual comment has the fields we need
    if "original" not in comment:
        log.debug(f"Comment missing 'original' field: {comment}.")
        return None

    original = comment.get("original", {})
    if not isinstance(original, dict):
        log.debug(f"Comment 'original' is not a dict, but is {type(original)}.")
        return None

    # Extract bits from original now that we've validated it
    text = original.get("text")
    author = original.get("author")
    initials = original.get("initials")

    if not text:
        log.debug("Comment has no text content.")
        return None

    if "reference_text" not in comment:
        log.debug(f"Comment is missing 'reference_text' field: {comment} Skipping.")
        return None

    if "id" not in comment:
        log.debug(f"Comment is missing 'id' field: {comment} Skipping.")
        return None

    return {
        "reference_text": comment["reference_text"],
        "id": comment["id"],
        "text": text,
        "author": author,
        "initials": initials,
    }


# endregion


# region safely_extract_heading_data
def safely_extract_heading_data(heading: dict) -> dict | None:
    """Extract heading data with validation."""
    if not isinstance(heading, dict):
        log.debug(f"Each stored heading should be a dict, got {type(heading)}")
        return None

    if "text" not in heading:
        log.debug(f"Heading missing required field 'text': {heading}")
        return None

    if "name" not in heading:
        log.debug(f"Heading missing required field 'name': {heading}")
        return None

    return {
        "text": heading["text"],
        "name": heading["name"],
        "style_id": heading.get(
            "style_id"
        ),  # Optional: if this field is available, add it. Otherwise its value is None.
    }


# endregion


# region safely_extract_experimental_formatting_data
def safely_extract_experimental_formatting_data(exp_fmt: dict) -> dict | None:
    """Extract experimental formatting data with validation."""
    if not isinstance(exp_fmt, dict):
        log.debug(f"Experimental formatting should be a dict, got {type(exp_fmt)}")
        return None

    if "ref_text" not in exp_fmt:
        log.debug(
            f"Experimental formatting missing required field: 'ref_text': {exp_fmt}"
        )
        return None

    if "formatting_type" not in exp_fmt:
        log.debug(
            f"Experimental formatting missing required field: 'formatting_type': {exp_fmt}"
        )
        return None

    return {
        "ref_text": exp_fmt["ref_text"],
        "formatting_type": exp_fmt["formatting_type"],
        "highlight_color_enum": exp_fmt.get(
            "highlight_color_enum"
        ),  # Optional field; will be set to None if not found
    }


# endregion
