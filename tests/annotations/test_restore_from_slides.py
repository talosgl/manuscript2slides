"""Tests for restore_from_slides module - JSON parsing, string manipulation, and range merging."""

import logging

import pytest

from manuscript2slides.annotations.restore_from_slides import (
    extract_slide_metadata,
    merge_overlapping_ranges,
    remove_ranges_from_text,
    safely_extract_comment_data,
    safely_extract_experimental_formatting_data,
    safely_extract_heading_data,
    split_speaker_notes,
)
from manuscript2slides.internals.constants import (
    METADATA_MARKER_FOOTER,
    METADATA_MARKER_HEADER,
    NOTES_MARKER_FOOTER,
    NOTES_MARKER_HEADER,
)
from manuscript2slides.models import SlideNotes

# region merge_overlapping_ranges tests


@pytest.mark.parametrize(
    "ranges,expected",
    [
        # Empty list
        ([], []),
        # Single range
        ([(0, 5)], [(0, 5)]),
        # Two non-overlapping ranges
        ([(0, 5), (10, 15)], [(0, 5), (10, 15)]),
        # Two overlapping ranges
        ([(0, 10), (5, 15)], [(0, 15)]),
        # Two adjacent ranges (touching)
        ([(0, 5), (5, 10)], [(0, 10)]),
        # Multiple overlapping ranges
        ([(0, 5), (3, 8), (7, 12)], [(0, 12)]),
        # Ranges with gaps
        ([(0, 5), (10, 15), (20, 25)], [(0, 5), (10, 15), (20, 25)]),
        # Unsorted ranges that overlap
        ([(10, 15), (0, 12)], [(0, 15)]),
        # Completely contained range
        ([(0, 20), (5, 10)], [(0, 20)]),
        # Multiple groups of overlapping ranges
        ([(0, 5), (3, 8), (20, 25), (23, 28)], [(0, 8), (20, 28)]),
    ],
)
def test_merge_overlapping_ranges(ranges: list, expected: list) -> None:
    """Test that overlapping ranges are correctly merged."""
    assert merge_overlapping_ranges(ranges) == expected


# endregion


# region remove_ranges_from_text tests


@pytest.mark.parametrize(
    "text,ranges,expected",
    [
        # No ranges
        ("Hello World", [], "Hello World"),
        # Single range
        ("Hello World", [(0, 6)], "World"),
        # Remove middle
        ("Hello World", [(5, 6)], "HelloWorld"),
        # Multiple non-overlapping ranges
        ("Hello World", [(0, 6), (10, 11)], "Worl"),
        # Overlapping ranges (should merge first)
        ("Hello World", [(0, 6), (5, 11)], ""),
        # Remove end
        ("Hello World", [(10, 11)], "Hello Worl"),
        # Remove everything
        ("Hello", [(0, 5)], ""),
    ],
)
def test_remove_ranges_from_text(text: str, ranges: list, expected: str) -> None:
    """Test that text ranges are correctly removed."""
    assert remove_ranges_from_text(text, ranges) == expected


# endregion


# region split_speaker_notes tests


def test_split_speaker_notes_empty_input() -> None:
    """Test with empty speaker notes.
    Case: if speaker notes are empty, then both user_notes and metadata should be empty."""
    result = split_speaker_notes("")
    assert isinstance(result, SlideNotes)
    assert result.user_notes == ""
    assert result.metadata == {}


def test_split_speaker_notes_only_user_notes() -> None:
    """Test with only user notes, no markers.
    Case: if there are only user notes (no markers), then notes should be populated, and metadata should be an empty dict."""
    notes = "These are my personal notes about the slide."
    result = split_speaker_notes(notes)
    assert result.user_notes == notes
    assert result.metadata == {}


def test_split_speaker_notes_with_valid_json() -> None:
    """Test extraction of valid JSON metadata.
    Case: if both user notes and JSON (with proper markers) exist, we preserve both, in the forms we expect."""
    json_content = '{"comments": [{"id": 1}], "footnotes": []}'
    notes = f"""User notes here.
{METADATA_MARKER_HEADER}
========
{json_content}
========
{METADATA_MARKER_FOOTER}"""

    result = split_speaker_notes(notes)
    assert result.user_notes == "User notes here."
    assert result.metadata == {"comments": [{"id": 1}], "footnotes": []}
    assert result.comments == [{"id": 1}]
    assert result.footnotes == []


def test_split_speaker_notes_with_invalid_json() -> None:
    """Test handling of invalid JSON in metadata section.
    Case: if invalid JSON is passed in (but with proper header/footer markers), we preserve the user notes,
    but still remove the JSON."""
    notes = f"""{METADATA_MARKER_HEADER}
{{invalid json}}
{METADATA_MARKER_FOOTER}
User notes."""

    result = split_speaker_notes(notes)
    assert result.user_notes == "User notes."
    assert result.metadata == {}


def test_split_speaker_notes_with_copied_notes_section() -> None:
    """Test that copied notes section is removed, and user notes are preserved."""
    notes = f"""User notes.
{NOTES_MARKER_HEADER}
Old copied annotations from previous run.
{NOTES_MARKER_FOOTER}"""

    result = split_speaker_notes(notes)
    assert result.user_notes == "User notes."
    assert NOTES_MARKER_HEADER not in result.user_notes


def test_split_speaker_notes_with_both_sections_and_user_notes_interspersed() -> None:
    """When the speaker notes contains all 3 things: json metadata, copied annotations, and user notes--
    AND the user notes are interspersed at the top, middle, and bottom of the marker sections, ensure we
    parse as expected; the user notes should all get into the user notes, json content gets assigned to
    the right place, and copied annotations should be removed."""
    json_content = '{"headings": [{"name": "Heading1", "text": "My Heading"}]}'
    notes = f"""My user notes at the top.

{METADATA_MARKER_HEADER}
========
{json_content}
========
{METADATA_MARKER_FOOTER}

User notes in the middle.

{NOTES_MARKER_HEADER}
Old copied text.
{NOTES_MARKER_FOOTER}

More user notes at the end."""

    result = split_speaker_notes(notes)
    assert METADATA_MARKER_HEADER not in result.user_notes
    assert NOTES_MARKER_HEADER not in result.user_notes
    assert "My user notes at the top." in result.user_notes
    assert "User notes in the middle." in result.user_notes
    assert "More user notes at the end." in result.user_notes
    assert "My Heading" not in result.user_notes
    assert result.headings == [{"name": "Heading1", "text": "My Heading"}]


def test_split_speaker_notes_missing_footer_markers() -> None:
    """Test when header exists but footer is missing. Case: Nothing should be removed; for malformed
    data we fallback to preservation."""
    notes = f"""{METADATA_MARKER_HEADER}
Some text but no footer.
User notes."""

    result = split_speaker_notes(notes)
    # Should not remove anything if markers are incomplete
    assert METADATA_MARKER_HEADER in result.user_notes
    assert "Some text but no footer." in result.user_notes
    assert "User notes." in result.user_notes


# endregion


# region extract_slide_metadata tests


def test_extract_slide_metadata_valid_dict() -> None:
    """Test extracting metadata from a valid dict."""
    metadata = {
        "comments": [{"id": 1, "text": "comment"}],
        "footnotes": [{"id": 1, "text": "footnote"}],
        "endnotes": [{"id": 1, "text": "endnote"}],
        "headings": [{"name": "H1", "text": "title"}],
        "experimental_formatting": [{"ref_text": "text", "formatting_type": "bold"}],
    }
    slide_notes = SlideNotes()
    result = extract_slide_metadata(metadata, slide_notes)

    assert result.metadata == metadata
    assert result.comments == metadata["comments"]
    assert result.footnotes == metadata["footnotes"]
    assert result.endnotes == metadata["endnotes"]
    assert result.headings == metadata["headings"]
    assert result.experimental_formatting == metadata["experimental_formatting"]


def test_extract_slide_metadata_non_dict_input(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that non-dict input returns unmutated SlideNotes."""
    slide_notes = SlideNotes()
    original_metadata = slide_notes.metadata.copy()  # Empty

    with caplog.at_level(logging.DEBUG):
        result = extract_slide_metadata("not a dict", slide_notes)  # type: ignore[arg-type]

    assert "should be a dict" in caplog.text
    assert result is slide_notes
    assert (
        result.metadata == original_metadata
    )  # It should still be empty, since invalid data was passed in


def test_extract_slide_metadata_invalid_list_types(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that non-list values for expected list fields are replaced with empty lists."""
    metadata = {
        "comments": "should be list",
        "footnotes": 123,
        "endnotes": {"key": "value"},
        "headings": None,
        "experimental_formatting": True,
    }
    slide_notes = SlideNotes()

    with caplog.at_level(logging.DEBUG):
        result = extract_slide_metadata(metadata, slide_notes)

    assert result.comments == []
    assert result.footnotes == []
    assert result.endnotes == []
    assert result.headings == []
    assert result.experimental_formatting == []

    # Verify all the type errors were logged
    assert "Comments from the slide notes JSON should be a list" in caplog.text
    assert "Footnotes from the slide notes JSON should be a list" in caplog.text
    assert "Endnotes from the slide notes JSON should be a list" in caplog.text
    assert "Headings from the slide notes JSON should be a list" in caplog.text
    assert (
        "Experimental_formatting from the slide notes JSON should be a list"
        in caplog.text
    )


def test_extract_slide_metadata_partial_fields() -> None:
    """Test with only some fields present."""
    metadata = {
        "comments": [{"id": 1}],
        # Missing footnotes, endnotes, headings, experimental_formatting
    }
    slide_notes = SlideNotes()
    result = extract_slide_metadata(metadata, slide_notes)

    assert result.comments == [{"id": 1}]
    assert result.footnotes == []
    assert result.endnotes == []
    assert result.headings == []
    assert result.experimental_formatting == []


# endregion


# region safely_extract_comment_data tests


def test_safely_extract_comment_data_valid() -> None:
    """Test extraction of valid comment data."""
    comment = {
        "id": "abc123",
        "reference_text": "highlighted text",
        "original": {
            "text": "This is the comment",
            "author": "Jane Eyre",
            "initials": "JE",
        },
    }
    result = safely_extract_comment_data(comment)

    assert result is not None
    assert result["id"] == "abc123"
    assert result["reference_text"] == "highlighted text"
    assert result["text"] == "This is the comment"
    assert result["author"] == "Jane Eyre"
    assert result["initials"] == "JE"


@pytest.mark.parametrize(
    "invalid_input",
    [
        "not a dict",
        [1, 2, 3],
        None,
    ],
)
def test_safely_extract_comment_data_non_dict(
    invalid_input,  # noqa: ANN001
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that non-dict input returns None."""
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(invalid_input)  # type: ignore[arg-type]

    assert result is None
    assert "should be a dict" in caplog.text


def test_safely_extract_comment_data_missing_original(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test comment missing 'original' field."""
    comment = {
        "id": "abc123",
        "reference_text": "text",
    }

    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(comment)

    assert result is None
    assert "missing 'original' field" in caplog.text


def test_safely_extract_comment_data_original_not_dict(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test comment where 'original' is not a dict."""
    comment = {
        "id": "abc123",
        "reference_text": "text",
        "original": "should be dict",
    }
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(comment)
    assert result is None
    assert "is not a dict" in caplog.text


def test_safely_extract_comment_data_missing_text(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test comment with original dict but no text."""
    comment = {
        "id": "abc123",
        "reference_text": "text",
        "original": {
            "author": "Jane",
        },
    }
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(comment)

    assert result is None
    assert "no text content" in caplog.text


def test_safely_extract_comment_data_empty_text(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test comment with empty text string."""
    comment = {
        "id": "abc123",
        "reference_text": "text",
        "original": {
            "text": "",
            "author": "Jane",
        },
    }
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(comment)

    assert result is None
    assert "no text content" in caplog.text


def test_safely_extract_comment_data_missing_reference_text(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test comment missing reference_text field."""
    comment = {
        "id": "abc123",
        "original": {
            "text": "comment text",
        },
    }
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(comment)

    assert result is None
    assert "missing 'reference_text' field" in caplog.text


def test_safely_extract_comment_data_missing_id(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test comment missing id field."""
    comment = {
        "reference_text": "text",
        "original": {
            "text": "comment text",
        },
    }
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_comment_data(comment)

    assert result is None
    assert "missing 'id' field" in caplog.text


def test_safely_extract_comment_data_optional_fields_missing() -> None:
    """Test that author and initials are optional."""
    comment = {
        "id": "abc123",
        "reference_text": "text",
        "original": {
            "text": "comment text",
        },
    }
    result = safely_extract_comment_data(comment)

    assert result is not None
    assert result["text"] == "comment text"
    assert result["author"] is None
    assert result["initials"] is None


# endregion


# region safely_extract_heading_data tests


def test_safely_extract_heading_data_valid() -> None:
    """Test extraction of valid heading data."""
    heading = {
        "text": "Introduction",
        "name": "Heading 1",
        "style_id": "Heading1",
    }
    result = safely_extract_heading_data(heading)

    assert result is not None
    assert result["text"] == "Introduction"
    assert result["name"] == "Heading 1"
    assert result["style_id"] == "Heading1"


@pytest.mark.parametrize(
    "invalid_input",
    [
        "not a dict",
        [1, 2, 3],
    ],
)
def test_safely_extract_heading_data_non_dict(
    invalid_input,  # noqa: ANN001
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that non-dict input returns None."""
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_heading_data(invalid_input)  # type: ignore[arg-type]

    assert result is None
    assert "should be a dict" in caplog.text


def test_safely_extract_heading_data_missing_text(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test heading missing required 'text' field."""
    heading = {"name": "Heading 1"}

    with caplog.at_level(logging.DEBUG):
        result = safely_extract_heading_data(heading)

    assert result is None
    assert "missing required field 'text'" in caplog.text


def test_safely_extract_heading_data_missing_name(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test heading missing required 'name' field."""
    heading = {"text": "Introduction"}

    with caplog.at_level(logging.DEBUG):
        result = safely_extract_heading_data(heading)

    assert result is None
    assert "missing required field 'name'" in caplog.text


def test_safely_extract_heading_data_optional_style_id_missing() -> None:
    """Test that style_id is optional."""
    heading = {
        "text": "Introduction",
        "name": "Heading 1",
    }
    result = safely_extract_heading_data(heading)

    assert result is not None
    assert result["text"] == "Introduction"
    assert result["name"] == "Heading 1"
    assert result["style_id"] is None


# endregion


# region safely_extract_experimental_formatting_data tests


def test_safely_extract_experimental_formatting_data_valid() -> None:
    """Test extraction of valid experimental formatting data."""
    exp_fmt = {
        "ref_text": "important text",
        "formatting_type": "highlight",
        "highlight_color_enum": "yellow",
    }
    result = safely_extract_experimental_formatting_data(exp_fmt)

    assert result is not None
    assert result["ref_text"] == "important text"
    assert result["formatting_type"] == "highlight"
    assert result["highlight_color_enum"] == "yellow"


@pytest.mark.parametrize(
    "invalid_input",
    [
        "not a dict",
        [1, 2, 3],
    ],
)
def test_safely_extract_experimental_formatting_data_non_dict(
    invalid_input,  # noqa: ANN001
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test that non-dict input returns None."""
    with caplog.at_level(logging.DEBUG):
        result = safely_extract_experimental_formatting_data(invalid_input)  # type: ignore[arg-type]

    assert result is None
    assert "should be a dict" in caplog.text


def test_safely_extract_experimental_formatting_data_missing_ref_text(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test formatting missing required 'ref_text' field."""
    exp_fmt = {"formatting_type": "highlight"}

    with caplog.at_level(logging.DEBUG):
        result = safely_extract_experimental_formatting_data(exp_fmt)

    assert result is None
    assert "missing required field: 'ref_text'" in caplog.text


def test_safely_extract_experimental_formatting_data_missing_formatting_type(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Test formatting missing required 'formatting_type' field."""
    exp_fmt = {"ref_text": "text"}

    with caplog.at_level(logging.DEBUG):
        result = safely_extract_experimental_formatting_data(exp_fmt)

    assert result is None
    assert "missing required field: 'formatting_type'" in caplog.text


def test_safely_extract_experimental_formatting_data_optional_color_missing() -> None:
    """Test that highlight_color_enum is optional."""
    exp_fmt = {
        "ref_text": "text",
        "formatting_type": "bold",
    }
    result = safely_extract_experimental_formatting_data(exp_fmt)

    assert result is not None
    assert result["ref_text"] == "text"
    assert result["formatting_type"] == "bold"
    assert result["highlight_color_enum"] is None


# endregion
