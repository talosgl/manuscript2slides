"""Tests to ensure we can create blank slide deck and blank document from our standard templates."""

# pyright: reportArgumentType=false

import logging
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

import manuscript2slides.templates as templates
from manuscript2slides.internals.define_config import UserConfig


# region test_create_empty_slide_deck
def test_create_empty_slide_deck_returns_empty_deck(
    path_to_sample_pptx_with_formatting: Path, path_to_sample_docx_with_formatting: Path
) -> None:
    """Ensure that the base presentation object is empty of slides,
    even if a pptx is passed in that template which contains slides."""
    # Arrange: Construct a cfg using a deck as template that we know has slides in it
    test_cfg = UserConfig(
        input_docx=path_to_sample_docx_with_formatting,
        template_pptx=path_to_sample_pptx_with_formatting,
    )

    # Action: Pass that to our function
    prs = templates.create_empty_slide_deck(test_cfg)

    # Assert the slide count is now 0
    assert len(prs.slides) == 0, f"prs.slides has {len(prs.slides)} slides in it."


@pytest.fixture
def path_to_missing_layout_pptx() -> Path:
    """Path to a pptx file that lacks the expected slide layout."""
    path = Path("tests/data/missing_layout.pptx")
    assert path.exists(), f"Test file not found: {path}"
    return path


def test_create_empty_slide_deck_fails_gracefully_without_slide_layout(
    path_to_missing_layout_pptx: Path,
    caplog: pytest.LogCaptureFixture,
    path_to_sample_docx_with_formatting: Path,
) -> None:
    """Ensure that if a pptx is passed in without the proper slide layout, we raise with a helpful message and log
    a helpful message to the error log."""

    # Arrange: Create a cfg whose template is pointing to a test fixture object that is missing our required slide layout

    test_cfg = UserConfig(
        input_docx=path_to_sample_docx_with_formatting,  # We pass a real docx so that the validation doesn't fail earlier
        template_pptx=path_to_missing_layout_pptx,
    )

    # Action: Capture logging at error level and pass our bad cfg into the func call
    with caplog.at_level(logging.ERROR):
        with pytest.raises(ValueError, match="layout"):
            templates.create_empty_slide_deck(test_cfg)

    # Assert: we should be catching the raise and we ought to be logging useful info
    assert "Could not find required slide layout" in caplog.text


# endregion


# region test_empty_document


def test_create_empty_document_returns_empty_doc(
    path_to_sample_docx_with_formatting: Path, path_to_sample_pptx_with_formatting: Path
) -> None:
    """Ensure that even if we're passed in a docx full of stuff as a template, we start with no pages/paragraphs."""
    # Arrange: Construct a cfg using a doc that we know has content in it
    test_cfg = UserConfig(
        input_pptx=path_to_sample_pptx_with_formatting,
        template_docx=path_to_sample_docx_with_formatting,
    )

    # Action: Pass that to our function
    docu = templates.create_empty_document(test_cfg)

    assert len(docu.paragraphs) == 0, (
        f"docu.paragraphs has {len(docu.paragraphs)} paragraphs in it."
    )


def test_create_empty_document_fails_gracefully_on_missing_styles(
    path_to_sample_pptx_with_formatting: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Ensure we raise and log a helpful message if a docx is provided as template that is missing the 'Normal' style."""
    # Arrange:
    # Create a Mock object to simulate a document.Document
    mock_doc = Mock()
    mock_style = Mock()
    mock_style.name = "Heading 1"
    mock_doc.styles = [mock_style]  # No "Normal" style

    # Action:
    with caplog.at_level(logging.ERROR):
        # Disguise our mock_doc as the return value for the docx.Document constructor
        with patch("docx.Document", return_value=mock_doc):
            test_cfg = UserConfig(input_pptx=path_to_sample_pptx_with_formatting)
            with pytest.raises(
                ValueError, match="styles"
            ):  # Test will fail if we do not raise as expected.
                templates.create_empty_document(test_cfg)

    # Assert that we also find helpful text in the error log
    assert "Could not find required default Word style.name 'Normal'" in caplog.text


# endregion
