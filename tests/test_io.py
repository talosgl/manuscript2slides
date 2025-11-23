"""Test I/O functions"""

# tests/test_io.py
import pytest
from pathlib import Path
from docx import document
from docx import Document
from manuscript2slides.io import load_and_validate_docx, validate_docx_path


# region test_validate_docx_path
def test_validate_docx_path_accepts_valid_file(
    sample_docx_with_formatting: Path,
) -> None:
    """Test that validate_docx_path accepts valid .docx files"""
    result = validate_docx_path(sample_docx_with_formatting)

    assert result == sample_docx_with_formatting
    assert result.exists()


def test_validate_docx_path_rejects_wrong_extension(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Test that non-.docx files are rejected"""
    wrong_file = tmp_path / "not_a_doc.txt"
    wrong_file.touch()

    with pytest.raises(ValueError, match="Expected a .docx"):
        validate_docx_path(wrong_file)

    assert "Wrong file extension" in caplog.text


def test_validate_docx_path_rejects_old_doc_format(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Test that old .doc format is rejected with helpful message"""
    old_doc = tmp_path / "old.doc"
    old_doc.touch()

    with pytest.raises(ValueError, match="only supports .docx files"):
        validate_docx_path(old_doc)

    assert "Unsupported .doc file" in caplog.text


# endregion


# region test_load_and_validate_docx
def test_load_docx_returns_non_empty_document_object(
    sample_docx_with_formatting: Path,
) -> None:
    """Test that loading a valid docx returns a Document object"""
    doc = load_and_validate_docx(sample_docx_with_formatting)

    assert isinstance(doc, document.Document)
    assert len(doc.paragraphs) > 0  # Should have some content


def test_load_docx_rejects_missing_file() -> None:
    """Test that a missing file passed in raises FileNotFoundError"""
    fake_path = Path("i_dont_exist.docx")

    with pytest.raises(FileNotFoundError, match="File not found"):
        load_and_validate_docx(fake_path)


def test_load_docx_rejects_empty_docx(tmp_path: Path) -> None:
    """Test that document with no content raises ValueError"""
    empty = tmp_path / "empty.docx"
    doc = Document()
    # Don't add any paragraphs
    doc.save(str(empty))

    with pytest.raises(ValueError, match="no paragraphs"):
        load_and_validate_docx(empty)

    doc.add_paragraph()
    doc.save(str(empty))

    with pytest.raises(ValueError, match="no text content"):
        load_and_validate_docx(empty)


# endregion
