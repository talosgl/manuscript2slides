"""Tests for standard formatting"""

# tests/test_formatting.py

# pyright: reportPrivateUsage=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportIndexIssue=false
# pyright: reportOptionalMemberAccess=false

from docx import Document
from pptx import Presentation
from docx import document
from pptx import presentation

from docx.text.run import Run as Run_docx
from docx.text.font import Font as Font_docx
from docx.text.paragraph import Paragraph as Paragraph_docx
from docx.shared import RGBColor as RGBColor_docx


from pptx.text.text import Font as Font_pptx
from pptx.text.text import _Paragraph as Paragraph_pptx
from pptx.text.text import _Run as Run_pptx
from pptx.dml.color import RGBColor as RGBColor_pptx

from pathlib import Path
import pytest
from manuscript2slides.processing import formatting
from manuscript2slides.internals.define_config import UserConfig

# region fixtures


@pytest.fixture
def blank_docx(path_to_empty_docx: Path) -> document.Document:
    """An empty docx object."""
    return Document(str(path_to_empty_docx))


@pytest.fixture
def pptx_with_formatting(
    path_to_sample_pptx_with_formatting: Path,
) -> presentation.Presentation:
    """A pptx.presentation.Presentation object with sample formatting."""
    return Presentation(path_to_sample_pptx_with_formatting)


@pytest.fixture
def pptx_w_twenty_empty_slides(
    path_to_pptx_w_twenty_empty_slides: Path,
) -> presentation.Presentation:
    """An empty pptx object that contains 20 empty slides so we don't need to add slides or slide layout stuff to be adding paragraphs and runs."""
    return Presentation(path_to_pptx_w_twenty_empty_slides)


@pytest.fixture
def docx_formatting_runs_dict(
    path_to_sample_docx_with_formatting: Path,
) -> dict[str, Run_docx]:
    """Pre-validated runs from test_formatting.docx for easy access."""

    docx_with_formatting = Document(str(path_to_sample_docx_with_formatting))

    test_err = "Test cannot proceed because test document does not match expectations; "

    # Bold, italics, underline
    bolded_run = docx_with_formatting.paragraphs[2].runs[0]
    assert bolded_run.font.bold is True, test_err + "expected bold run at para 2, run 0"

    italic_run = docx_with_formatting.paragraphs[3].runs[0]
    assert italic_run.font.italic is True, (
        test_err + "expected italic in run at para 3, run 0"
    )

    bold_and_italic_run = docx_with_formatting.paragraphs[4].runs[0]
    assert (
        bold_and_italic_run.font.bold is True
        and bold_and_italic_run.font.italic is True
    ), (test_err + "expected italic + bold at para 4, run 0")

    underline_run = docx_with_formatting.paragraphs[5].runs[0]
    assert underline_run.font.underline is True, (
        test_err + "expected underline at para 5, run 0"
    )

    bold_and_underline_run = docx_with_formatting.paragraphs[6].runs[0]
    assert (
        bold_and_underline_run.font.bold is True
        and bold_and_underline_run.font.underline is True
    ), (test_err + "expected bold + underline at para 6, run 0")

    underline_and_italic_run = docx_with_formatting.paragraphs[7].runs[0]
    assert (
        underline_and_italic_run.font.underline is True
        and underline_and_italic_run.font.italic is True
    ), (test_err + "expected underline + italic at para 7, run 0")

    all_three_run = docx_with_formatting.paragraphs[8].runs[0]
    assert (
        all_three_run.font.bold is True
        and all_three_run.font.italic is True
        and all_three_run.font.underline is True
    ), (test_err + "expected bold + italic + underline at para 8, run 0")

    # font.name or Typeface
    typeface_run = docx_with_formatting.paragraphs[9].runs[0]
    assert typeface_run.font.name is not None and typeface_run.font.name == "Georgia", (
        test_err + "expected font.name to be set to 'Georgia' at para 9, run 0"
    )

    # Size
    large_size_run = docx_with_formatting.paragraphs[11].runs[0]
    assert large_size_run.font.size.pt == 48

    small_size_run = docx_with_formatting.paragraphs[12].runs[0]
    assert small_size_run.font.size.pt == 8

    # Color
    color_para = docx_with_formatting.paragraphs[14]
    red_run = color_para.runs[1]
    assert red_run.font.color.rgb is not None, (
        test_err + "expected font.color.rgb to be set at para 14, run 1"
    )
    assert red_run.font.color.rgb == RGBColor_docx(0xFF, 0x00, 0x00), (
        test_err
        + "expected font.color.rgb to be red (RGBColor_docx(0xFF, 0x00, 0x00)) at para 14, run 1"
    )

    blue_run = color_para.runs[3]
    assert blue_run.font.color.rgb is not None, (
        test_err + "expected font.color.rgb at para 14, run 3"
    )
    assert blue_run.font.color.rgb == RGBColor_docx(0x44, 0x72, 0xC4), (
        test_err
        + "expected font.color.rgb to be blue RGBColor_docx(0x44, 0x72, 0xC4)) at para 14, run 1"
    )

    # Experimental formatting
    hl_run = docx_with_formatting.paragraphs[34].runs[0]
    assert hl_run.font.highlight_color is not None, (
        test_err + "expected highlight_color at para 34, run 0"
    )

    complex_hl_para = docx_with_formatting.paragraphs[35]
    assert "highlighted, but also" in complex_hl_para.text, (
        test_err + "expected text 'highlighted, but also' in para 35"
    )
    hl_run2 = complex_hl_para.runs[0]
    assert hl_run2.font.highlight_color is not None, (
        test_err + "expected highlight_color at para 35, run 0"
    )

    hl_and_underlined = complex_hl_para.runs[1]
    assert (
        hl_and_underlined.font.highlight_color is not None
        and hl_and_underlined.font.underline is True
    ), (test_err + "expected highlight + underline at para 35, run 1")

    hl_only = complex_hl_para.runs[2]
    assert hl_only.font.highlight_color is not None, (
        test_err + "expected highlight_color at para 35, run 2"
    )

    hl_and_italic = complex_hl_para.runs[3]
    assert (
        hl_and_italic.font.highlight_color is not None
        and hl_and_italic.font.italic is True
    ), (test_err + "expected highlight + italic at para 35, run 3")

    hl_and_bold = complex_hl_para.runs[5]
    assert (
        hl_and_bold.font.highlight_color is not None and hl_and_bold.font.bold is True
    ), (test_err + "expected highlight + bold at para 35, run 5")

    super_sub_para = docx_with_formatting.paragraphs[36]
    assert "superscript" in super_sub_para.text, (
        test_err + "expected text 'superscript' in para 36"
    )
    subscript_run = super_sub_para.runs[1]
    assert subscript_run.font.subscript is True, (
        test_err + "expected subscript at para 36, run 1"
    )

    superscript_run = super_sub_para.runs[3]
    assert superscript_run.font.superscript is True, (
        test_err + "expected superscript at para 36, run 3"
    )

    all_caps_run = docx_with_formatting.paragraphs[37].runs[0]
    assert all_caps_run.font.all_caps is True, (
        test_err + "expected all_caps at para 37, run 0"
    )

    small_caps_run = docx_with_formatting.paragraphs[38].runs[0]
    assert small_caps_run.font.small_caps is True, (
        test_err + "expected small_caps at para 38, run 0"
    )

    strike_para = docx_with_formatting.paragraphs[39]
    single_str_run = strike_para.runs[1]
    assert single_str_run.font.strike is True, (
        test_err + "expected strike at para 39, run 1"
    )

    dbl_str_run = strike_para.runs[3]
    assert dbl_str_run.font.double_strike is True, (
        test_err + "expected double_strike at para 39, run 3"
    )

    return {
        "bold": bolded_run,
        "italic": italic_run,
        "bold_and_italic": bold_and_italic_run,
        "underline": underline_run,
        "bold_and_underline": bold_and_underline_run,
        "underline_and_italic": underline_and_italic_run,
        "all_three": all_three_run,
        "typeface": typeface_run,
        "red_run": red_run,
        "blue_run": blue_run,
        "large_size": large_size_run,
        "small_size": small_size_run,
        "highlight": hl_run,
        "highlight_2": hl_run2,
        "highlight_and_underline": hl_and_underlined,
        "highlight_only": hl_only,
        "highlight_and_italic": hl_and_italic,
        "highlight_and_bold": hl_and_bold,
        "subscript": subscript_run,
        "superscript": superscript_run,
        "all_caps": all_caps_run,
        "small_caps": small_caps_run,
        "single_strike": single_str_run,
        "double_strike": dbl_str_run,
    }


@pytest.fixture
def pptx_formatting_runs_dict(
    path_to_sample_pptx_with_formatting: Path,
) -> dict[str, Run_pptx]:
    """Pre-validated runs from test_formatting_expected_output.pptx for easy access."""
    pptx_with_formatting = Presentation(path_to_sample_pptx_with_formatting)
    test_err = "Test cannot proceed because test document does not match expectations; "

    bold_slide = pptx_with_formatting.slides[1]
    bold_paragraphs = bold_slide.shapes.placeholders[1].text_frame.paragraphs
    bold_run = bold_paragraphs[0].runs[0]
    assert bold_run.font.bold == True, (
        test_err + "expected bold run at slide 1, placeholder 1, para 0, run 0"
    )

    italic_run = (
        pptx_with_formatting.slides[2]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert italic_run.font.italic is True, (
        test_err + "expected italic run at slide 2, placeholder 1, para 0, run 0"
    )

    bold_and_italic_run = (
        pptx_with_formatting.slides[3]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert (
        bold_and_italic_run.font.bold is True
        and bold_and_italic_run.font.italic is True
    ), (test_err + "expected bold + italic at slide 3, placeholder 1, para 0, run 0")

    underline_run = (
        pptx_with_formatting.slides[4]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert underline_run.font.underline is True, (
        test_err + "expected underline at slide 4, placeholder 1, para 0, run 0"
    )

    bold_and_underline_run = (
        pptx_with_formatting.slides[5]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert (
        bold_and_underline_run.font.bold is True
        and bold_and_underline_run.font.underline is True
    ), (test_err + "expected bold + underline at slide 5, placeholder 1, para 0, run 0")

    underline_and_italic_run = (
        pptx_with_formatting.slides[6]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert (
        underline_and_italic_run.font.underline is True
        and underline_and_italic_run.font.italic is True
    ), (
        test_err
        + "expected underline + italic at slide 6, placeholder 1, para 0, run 0"
    )

    all_three_run = (
        pptx_with_formatting.slides[7]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert (
        all_three_run.font.bold is True
        and all_three_run.font.italic is True
        and all_three_run.font.underline is True
    ), (
        test_err
        + "expected bold + italic + underline at slide 7, placeholder 1, para 0, run 0"
    )

    typeface_run = (
        pptx_with_formatting.slides[8]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert typeface_run.font.name is not None and typeface_run.font.name == "Georgia", (
        test_err
        + "expected font.name to be 'Georgia' at slide 8, placeholder 1, para 0, run 0"
    )

    # skip 9

    large_size_run = (
        pptx_with_formatting.slides[10]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert large_size_run.font.size.pt == 48, (
        test_err + "expected font size 48pt at slide 10, placeholder 1, para 0, run 0"
    )

    small_size_run = (
        pptx_with_formatting.slides[11]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert small_size_run.font.size.pt == 8, (
        test_err + "expected font size 8pt at slide 11, placeholder 1, para 0, run 0"
    )

    # skip 12

    # Color - assuming similar to docx where there are multiple runs
    color_para = (
        pptx_with_formatting.slides[13].shapes.placeholders[1].text_frame.paragraphs[0]
    )
    red_run = color_para.runs[1]  # guessing run 1 for red
    assert red_run.font.color.rgb is not None and red_run.font.color.rgb == (
        255,
        0,
        0,
    ), (
        test_err + "expected red color at slide 13, placeholder 1, para 0, run 1"
    )
    blue_run = color_para.runs[3]  # guessing run 3 for blue
    assert blue_run.font.color.rgb is not None and blue_run.font.color.rgb == (
        68,
        114,
        196,
    ), (
        test_err + "expected blue color at slide 13, placeholder 1, para 0, run 3"
    )

    # skip 14

    # Experimental formatting - using XML assertions
    hl_run = (
        pptx_with_formatting.slides[15]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    # Check for highlight in XML since pptx doesn't expose as property
    hl_xml = hl_run._r.xml
    assert "a:highlight" in hl_xml, (
        test_err
        + "expected highlight element in XML at slide 15, placeholder 1, para 0, run 0"
    )

    # Complex highlight paragraph
    hl_para = (
        pptx_with_formatting.slides[16].shapes.placeholders[1].text_frame.paragraphs[0]
    )
    # Get the highlighted runs from this paragraph
    hl_run2 = hl_para.runs[0]
    assert "a:highlight" in hl_run2._r.xml, (
        test_err + "expected highlight in run 0 at slide 16"
    )

    hl_and_underlined = hl_para.runs[1]
    assert (
        "a:highlight" in hl_and_underlined._r.xml
        and hl_and_underlined.font.underline is True
    ), (test_err + "expected highlight + underline at slide 16, run 1")

    super_sub_para = (
        pptx_with_formatting.slides[17].shapes.placeholders[1].text_frame.paragraphs[0]
    )
    subscript_run = super_sub_para.runs[1]
    baseline_sub = subscript_run.font._element.get("baseline")
    assert baseline_sub is not None and int(baseline_sub) < 0, (
        test_err + "expected negative baseline for subscript at slide 17, para 0, run 1"
    )

    superscript_run = super_sub_para.runs[3]
    baseline_super = superscript_run.font._element.get("baseline")
    assert baseline_super is not None and int(baseline_super) > 0, (
        test_err
        + "expected positive baseline for superscript at slide 17, para 0, run 3"
    )

    # All caps
    all_caps_run = (
        pptx_with_formatting.slides[18]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert all_caps_run.font._element.get("cap") == "all", (
        test_err
        + "expected cap='all' attribute at slide 18, placeholder 1, para 0, run 0"
    )

    # Small caps
    small_caps_run = (
        pptx_with_formatting.slides[19]
        .shapes.placeholders[1]
        .text_frame.paragraphs[0]
        .runs[0]
    )
    assert small_caps_run.font._element.get("cap") == "small", (
        test_err
        + "expected cap='small' attribute at slide 19, placeholder 1, para 0, run 0"
    )

    # Strike paragraph
    strike_para = (
        pptx_with_formatting.slides[20].shapes.placeholders[1].text_frame.paragraphs[0]
    )
    single_str_run = strike_para.runs[1]
    assert single_str_run.font._element.get("strike") == "sngStrike", (
        test_err + "expected strike='sngStrike' attribute at slide 20, para 0, run 1"
    )

    dbl_str_run = strike_para.runs[3]
    assert dbl_str_run.font._element.get("strike") == "dblStrike", (
        test_err + "expected strike='dblStrike' attribute at slide 20, para 0, run 3"
    )

    return {
        "bold": bold_run,
        "italic": italic_run,
        "bold_and_italic": bold_and_italic_run,
        "underline": underline_run,
        "bold_and_underline": bold_and_underline_run,
        "underline_and_italic": underline_and_italic_run,
        "all_three": all_three_run,
        "typeface": typeface_run,
        "large_size": large_size_run,
        "small_size": small_size_run,
        "red_run": red_run,
        "blue_run": blue_run,
        "highlight": hl_run,
        "highlight_2": hl_run2,
        "highlight_and_underline": hl_and_underlined,
        "subscript": subscript_run,
        "superscript": superscript_run,
        "all_caps": all_caps_run,
        "small_caps": small_caps_run,
        "single_strike": single_str_run,
        "double_strike": dbl_str_run,
    }


# endregion


# region test helpers


def _create_target_pptx_run(pptx_obj: presentation.Presentation) -> Run_pptx:
    """Helper to create a target pptx run for testing."""
    slide = pptx_obj.slides[0]
    text_frame = slide.shapes.placeholders[1].text_frame
    target_para = text_frame.paragraphs[0]
    target_run = target_para.add_run()
    target_run.text = "Test pptx Run"
    return target_run


def _create_target_docx_run(docx_obj: document.Document) -> Run_docx:
    """Helper to create a target docx run for testing."""
    paragraph = docx_obj.add_paragraph()
    target_run = paragraph.add_run("Test docx Run. ")
    return target_run


# endregion


# region docx2pptx RUN tests


# region _copy_basic_font_formatting tests


def test_copy_basic_font_formatting_preserves_typeface(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies font.name typeface."""
    source_run = docx_formatting_runs_dict["typeface"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.name is not None and target_run.font.name == "Georgia"


def test_copy_basic_font_formatting_preserves_bold(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold formatting."""
    source_run = docx_formatting_runs_dict["bold"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True


def test_copy_basic_font_formatting_preserves_italic(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies italic formatting."""
    source_run = docx_formatting_runs_dict["italic"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.italic is True


def test_copy_basic_font_formatting_preserves_underline(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies underline formatting."""
    source_run = docx_formatting_runs_dict["underline"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.underline is True


def test_copy_basic_font_formatting_preserves_bold_and_italic(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold and italic formatting together."""
    source_run = docx_formatting_runs_dict["bold_and_italic"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True
    assert target_run.font.italic is True


def test_copy_basic_font_formatting_preserves_bold_and_underline(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold and underline formatting together."""
    source_run = docx_formatting_runs_dict["bold_and_underline"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True
    assert target_run.font.underline is True


def test_copy_basic_font_formatting_preserves_underline_and_italic(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies underline and italic formatting together."""
    source_run = docx_formatting_runs_dict["underline_and_italic"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.underline is True
    assert target_run.font.italic is True


def test_copy_basic_font_formatting_preserves_all_three(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold, italic, and underline formatting together."""
    source_run = docx_formatting_runs_dict["all_three"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True
    assert target_run.font.italic is True
    assert target_run.font.underline is True


# endregion

# region _copy_font_size_formatting tests


def test_copy_font_size_formatting_preserves_large_size(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_font_size_formatting correctly copies large font size (48pt)."""
    source_run = docx_formatting_runs_dict["large_size"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_font_size_formatting(source_run.font, target_run.font)

    assert target_run.font.size.pt == 48


def test_copy_font_size_formatting_preserves_small_size(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_font_size_formatting correctly copies small font size (8pt)."""
    source_run = docx_formatting_runs_dict["small_size"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_font_size_formatting(source_run.font, target_run.font)

    assert target_run.font.size.pt == 8


# endregion

# region _copy_font_color_formatting tests


def test_copy_font_color_formatting_preserves_color(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that _copy_font_color_formatting correctly copies colors."""
    source_run = docx_formatting_runs_dict["red_run"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting._copy_font_color_formatting(source_run.font, target_run.font)

    assert target_run.font.color.rgb is not None and target_run.font.color.rgb == (
        255,
        0,
        0,
    )


# endregion

# region _copy_experimental_formatting_docx2pptx tests


def test_copy_experimental_formatting_docx2pptx_preserves_highlight(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies highlight from docx to pptx."""
    source_run = docx_formatting_runs_dict["highlight"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML contains highlight
    assert "a:highlight" in target_run._r.xml
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "highlight"


def test_copy_experimental_formatting_docx2pptx_preserves_single_strike(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies single strike from docx to pptx."""
    source_run = docx_formatting_runs_dict["single_strike"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML attribute
    assert target_run.font._element.get("strike") == "sngStrike"
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "strike"


def test_copy_experimental_formatting_docx2pptx_preserves_double_strike(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies double strike from docx to pptx."""
    source_run = docx_formatting_runs_dict["double_strike"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML attribute
    assert target_run.font._element.get("strike") == "dblStrike"
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "double_strike"


def test_copy_experimental_formatting_docx2pptx_preserves_subscript(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies subscript from docx to pptx."""
    source_run = docx_formatting_runs_dict["subscript"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML attribute (negative baseline)
    baseline = target_run.font._element.get("baseline")
    assert baseline is not None and int(baseline) < 0
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "subscript"


def test_copy_experimental_formatting_docx2pptx_preserves_superscript(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies superscript from docx to pptx."""
    source_run = docx_formatting_runs_dict["superscript"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML attribute (positive baseline)
    baseline = target_run.font._element.get("baseline")
    assert baseline is not None and int(baseline) > 0
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "superscript"


def test_copy_experimental_formatting_docx2pptx_preserves_all_caps(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies all caps from docx to pptx."""
    source_run = docx_formatting_runs_dict["all_caps"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML attribute
    assert target_run.font._element.get("cap") == "all"
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "all_caps"


def test_copy_experimental_formatting_docx2pptx_preserves_small_caps(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that experimental formatting correctly copies small caps from docx to pptx."""
    source_run = docx_formatting_runs_dict["small_caps"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    metadata = []

    formatting._copy_experimental_formatting_docx2pptx(source_run, target_run, metadata)

    # Check XML attribute
    assert target_run.font._element.get("cap") == "small"
    # Check metadata was recorded
    assert len(metadata) == 1
    assert metadata[0]["formatting_type"] == "small_caps"


# endregion

# region copy_run_formatting_docx2pptx tests basics


def test_copy_run_formatting_docx2pptx_copies_all_basic_formatting(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that copy_run_formatting_docx2pptx copies typeface, bold, italic, and underline formatting attributes."""

    # Use a run with multiple formatting attributes
    source_run = docx_formatting_runs_dict["all_three"]  # bold + italic + underline
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting.copy_run_formatting_docx2pptx(
        source_run, target_run, [], UserConfig(experimental_formatting_on=False)
    )

    assert target_run.font.bold is True
    assert target_run.font.italic is True
    assert target_run.font.underline is True


def test_copy_run_formatting_docx2pptx_copies_typeface(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that orchestrator copy_run_formatting_docx2pptx properly copies typeface."""

    source_run = docx_formatting_runs_dict["typeface"]
    target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)

    formatting.copy_run_formatting_docx2pptx(
        source_run, target_run, [], UserConfig(experimental_formatting_on=False)
    )

    assert target_run.font.name is not None and target_run.font.name == "Georgia"


def test_copy_run_formatting_docx2pptx_copies_size_formatting(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that orchestrator copy_run_formatting_docx2pptx properly copies size."""

    small_size_source_run = docx_formatting_runs_dict["small_size"]
    small_size_target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    formatting.copy_run_formatting_docx2pptx(
        small_size_source_run,
        small_size_target_run,
        [],
        UserConfig(experimental_formatting_on=False),
    )

    large_size_source_run = docx_formatting_runs_dict["large_size"]
    large_size_target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    formatting.copy_run_formatting_docx2pptx(
        large_size_source_run,
        large_size_target_run,
        [],
        UserConfig(experimental_formatting_on=False),
    )

    assert small_size_target_run.font.size.pt == 8
    assert large_size_target_run.font.size.pt == 48


def test_copy_run_formatting_docx2pptx_copies_color_formatting(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that orchestrator copy_run_formatting_docx2pptx properly copies color."""

    red_source_run = docx_formatting_runs_dict["red_run"]
    red_target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    formatting.copy_run_formatting_docx2pptx(
        red_source_run,
        red_target_run,
        [],
        UserConfig(experimental_formatting_on=False),
    )

    blue_source_run = docx_formatting_runs_dict["blue_run"]
    blue_target_run = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    formatting.copy_run_formatting_docx2pptx(
        blue_source_run,
        blue_target_run,
        [],
        UserConfig(experimental_formatting_on=False),
    )

    assert (
        red_target_run.font.color.rgb is not None
        and red_target_run.font.color.rgb == (255, 0, 0)
    )
    assert (
        blue_target_run.font.color.rgb is not None
        and blue_target_run.font.color.rgb == (68, 114, 196)
    )


# endregion

# region copy_run_formatting_docx2pptx experimental tests


def test_copy_run_formatting_docx2pptx_copies_experimental_when_enabled(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that orchestrator copy_run_formatting_docx2pptx copies experimental formatting when enabled."""
    # Test with highlight
    hl_source = docx_formatting_runs_dict["highlight"]
    hl_target = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    hl_metadata = []
    formatting.copy_run_formatting_docx2pptx(
        hl_source, hl_target, hl_metadata, UserConfig(experimental_formatting_on=True)
    )
    assert "a:highlight" in hl_target._r.xml
    assert len(hl_metadata) == 1

    # Test with subscript
    sub_source = docx_formatting_runs_dict["subscript"]
    sub_target = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    sub_metadata = []
    formatting.copy_run_formatting_docx2pptx(
        sub_source, sub_target, sub_metadata, UserConfig(experimental_formatting_on=True)
    )
    baseline = sub_target.font._element.get("baseline")
    assert baseline is not None and int(baseline) < 0
    assert len(sub_metadata) == 1

    # Test with all caps
    caps_source = docx_formatting_runs_dict["all_caps"]
    caps_target = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    caps_metadata = []
    formatting.copy_run_formatting_docx2pptx(
        caps_source, caps_target, caps_metadata, UserConfig(experimental_formatting_on=True)
    )
    assert caps_target.font._element.get("cap") == "all"
    assert len(caps_metadata) == 1


def test_copy_run_formatting_docx2pptx_skips_experimental_when_disabled(
    docx_formatting_runs_dict: dict[str, Run_docx],
    pptx_w_twenty_empty_slides: presentation.Presentation,
) -> None:
    """Test that orchestrator copy_run_formatting_docx2pptx skips experimental formatting when disabled."""
    # Test with highlight
    hl_source = docx_formatting_runs_dict["highlight"]
    hl_target = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    hl_metadata = []
    formatting.copy_run_formatting_docx2pptx(
        hl_source, hl_target, hl_metadata, UserConfig(experimental_formatting_on=False)
    )
    assert "a:highlight" not in hl_target._r.xml
    assert len(hl_metadata) == 0

    # Test with subscript
    sub_source = docx_formatting_runs_dict["subscript"]
    sub_target = _create_target_pptx_run(pptx_w_twenty_empty_slides)
    sub_metadata = []
    formatting.copy_run_formatting_docx2pptx(
        sub_source, sub_target, sub_metadata, UserConfig(experimental_formatting_on=False)
    )
    baseline = sub_target.font._element.get("baseline")
    assert baseline is None
    assert len(sub_metadata) == 0


# endregion


# endregion docx2pptx RUN tests

# region TODO: docx2pptx paragraph formatting

# region TODO: get_effective_font_name_docx
# endregion

# region TODO: copy_paragraph_formatting_docx2pptx
# endregion

# region TODO: _copy_paragraph_font_name_docx2pptx
# endregion

# region TODO: _copy_paragraph_alignment_docx2pptx
# endregion

# endregion


# region pptx2docx RUN tests


# region _copy_basic_font_formatting tests (pptx2docx)


def test_copy_basic_font_formatting_pptx2docx_preserves_typeface(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies font.name typeface from pptx to docx."""
    source_run = pptx_formatting_runs_dict["typeface"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.name is not None and target_run.font.name == "Georgia"


def test_copy_basic_font_formatting_pptx2docx_preserves_bold(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold formatting from pptx to docx."""
    source_run = pptx_formatting_runs_dict["bold"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True


def test_copy_basic_font_formatting_pptx2docx_preserves_italic(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies italic formatting from pptx to docx."""
    source_run = pptx_formatting_runs_dict["italic"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.italic is True


def test_copy_basic_font_formatting_pptx2docx_preserves_underline(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies underline formatting from pptx to docx."""
    source_run = pptx_formatting_runs_dict["underline"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.underline is True


def test_copy_basic_font_formatting_pptx2docx_preserves_bold_and_italic(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold and italic formatting together from pptx to docx."""
    source_run = pptx_formatting_runs_dict["bold_and_italic"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True
    assert target_run.font.italic is True


def test_copy_basic_font_formatting_pptx2docx_preserves_bold_and_underline(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold and underline formatting together from pptx to docx."""
    source_run = pptx_formatting_runs_dict["bold_and_underline"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True
    assert target_run.font.underline is True


def test_copy_basic_font_formatting_pptx2docx_preserves_underline_and_italic(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies underline and italic formatting together from pptx to docx."""
    source_run = pptx_formatting_runs_dict["underline_and_italic"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.underline is True
    assert target_run.font.italic is True


def test_copy_basic_font_formatting_pptx2docx_preserves_all_three(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_basic_font_formatting correctly copies bold, italic, and underline formatting together from pptx to docx."""
    source_run = pptx_formatting_runs_dict["all_three"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_basic_font_formatting(source_run.font, target_run.font)

    assert target_run.font.bold is True
    assert target_run.font.italic is True
    assert target_run.font.underline is True


# endregion

# region _copy_font_size_formatting tests (pptx2docx)


def test_copy_font_size_formatting_pptx2docx_preserves_large_size(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_font_size_formatting correctly copies large font size (48pt) from pptx to docx."""
    source_run = pptx_formatting_runs_dict["large_size"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_font_size_formatting(source_run.font, target_run.font)

    assert target_run.font.size.pt == 48


def test_copy_font_size_formatting_pptx2docx_preserves_small_size(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_font_size_formatting correctly copies small font size (8pt) from pptx to docx."""
    source_run = pptx_formatting_runs_dict["small_size"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_font_size_formatting(source_run.font, target_run.font)

    assert target_run.font.size.pt == 8


# endregion

# region _copy_font_color_formatting tests (pptx2docx)


def test_copy_font_color_formatting_pptx2docx_preserves_color(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that _copy_font_color_formatting correctly copies colors from pptx to docx."""
    source_run = pptx_formatting_runs_dict["red_run"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_font_color_formatting(source_run.font, target_run.font)

    assert target_run.font.color.rgb is not None and target_run.font.color.rgb == (
        255,
        0,
        0,
    )


# endregion

# region copy_run_formatting_pptx2docx tests basics


def test_copy_run_formatting_pptx2docx_copies_all_basic_formatting(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that copy_run_formatting_pptx2docx copies typeface, bold, italic, and underline formatting attributes."""

    # Use a run with multiple formatting attributes
    source_run = pptx_formatting_runs_dict["all_three"]  # bold + italic + underline
    target_run = _create_target_docx_run(blank_docx)

    formatting.copy_run_formatting_pptx2docx(
        source_run, target_run, UserConfig(experimental_formatting_on=False)
    )

    assert target_run.font.bold is True
    assert target_run.font.italic is True
    assert target_run.font.underline is True


def test_copy_run_formatting_pptx2docx_copies_typeface(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that orchestrator copy_run_formatting_pptx2docx properly copies typeface."""

    source_run = pptx_formatting_runs_dict["typeface"]
    target_run = _create_target_docx_run(blank_docx)

    formatting.copy_run_formatting_pptx2docx(
        source_run, target_run, UserConfig(experimental_formatting_on=False)
    )

    assert target_run.font.name is not None and target_run.font.name == "Georgia"


def test_copy_run_formatting_pptx2docx_copies_size_formatting(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that orchestrator copy_run_formatting_pptx2docx properly copies size."""

    small_size_source_run = pptx_formatting_runs_dict["small_size"]
    small_size_target_run = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        small_size_source_run,
        small_size_target_run,
        UserConfig(experimental_formatting_on=False),
    )

    large_size_source_run = pptx_formatting_runs_dict["large_size"]
    large_size_target_run = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        large_size_source_run,
        large_size_target_run,
        UserConfig(experimental_formatting_on=False),
    )

    assert small_size_target_run.font.size.pt == 8
    assert large_size_target_run.font.size.pt == 48


def test_copy_run_formatting_pptx2docx_copies_color_formatting(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that orchestrator copy_run_formatting_pptx2docx properly copies color."""

    red_source_run = pptx_formatting_runs_dict["red_run"]
    red_target_run = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        red_source_run,
        red_target_run,
        UserConfig(experimental_formatting_on=False),
    )

    blue_source_run = pptx_formatting_runs_dict["blue_run"]
    blue_target_run = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        blue_source_run,
        blue_target_run,
        UserConfig(experimental_formatting_on=False),
    )

    assert (
        red_target_run.font.color.rgb is not None
        and red_target_run.font.color.rgb == (255, 0, 0)
    )
    assert (
        blue_target_run.font.color.rgb is not None
        and blue_target_run.font.color.rgb == (68, 114, 196)
    )


# endregion

# region _copy_experimental_formatting_pptx2docx tests


def test_copy_experimental_formatting_pptx2docx_preserves_highlight(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies highlight from pptx to docx."""
    source_run = pptx_formatting_runs_dict["highlight"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.highlight_color is not None


def test_copy_experimental_formatting_pptx2docx_preserves_single_strike(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies single strike from pptx to docx."""
    source_run = pptx_formatting_runs_dict["single_strike"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.strike is True


def test_copy_experimental_formatting_pptx2docx_preserves_double_strike(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies double strike from pptx to docx."""
    source_run = pptx_formatting_runs_dict["double_strike"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.double_strike is True


def test_copy_experimental_formatting_pptx2docx_preserves_subscript(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies subscript from pptx to docx."""
    source_run = pptx_formatting_runs_dict["subscript"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.subscript is True


def test_copy_experimental_formatting_pptx2docx_preserves_superscript(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies superscript from pptx to docx."""
    source_run = pptx_formatting_runs_dict["superscript"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.superscript is True


def test_copy_experimental_formatting_pptx2docx_preserves_all_caps(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies all caps from pptx to docx."""
    source_run = pptx_formatting_runs_dict["all_caps"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.all_caps is True


def test_copy_experimental_formatting_pptx2docx_preserves_small_caps(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that experimental formatting correctly copies small caps from pptx to docx."""
    source_run = pptx_formatting_runs_dict["small_caps"]
    target_run = _create_target_docx_run(blank_docx)

    formatting._copy_experimental_formatting_pptx2docx(source_run, target_run)

    assert target_run.font.small_caps is True


# endregion

# region copy_run_formatting_pptx2docx experimental tests


def test_copy_run_formatting_pptx2docx_copies_experimental_when_enabled(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that orchestrator copy_run_formatting_pptx2docx copies experimental formatting when enabled."""
    # Test with highlight
    hl_source = pptx_formatting_runs_dict["highlight"]
    hl_target = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        hl_source, hl_target, UserConfig(experimental_formatting_on=True)
    )
    assert hl_target.font.highlight_color is not None

    # Test with subscript
    sub_source = pptx_formatting_runs_dict["subscript"]
    sub_target = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        sub_source, sub_target, UserConfig(experimental_formatting_on=True)
    )
    assert sub_target.font.subscript is True

    # Test with all caps
    caps_source = pptx_formatting_runs_dict["all_caps"]
    caps_target = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        caps_source, caps_target, UserConfig(experimental_formatting_on=True)
    )
    assert caps_target.font.all_caps is True


def test_copy_run_formatting_pptx2docx_skips_experimental_when_disabled(
    pptx_formatting_runs_dict: dict[str, Run_pptx],
    blank_docx: document.Document,
) -> None:
    """Test that orchestrator copy_run_formatting_pptx2docx skips experimental formatting when disabled."""
    # Test with highlight
    hl_source = pptx_formatting_runs_dict["highlight"]
    hl_target = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        hl_source, hl_target, UserConfig(experimental_formatting_on=False)
    )
    assert hl_target.font.highlight_color is None

    # Test with subscript
    sub_source = pptx_formatting_runs_dict["subscript"]
    sub_target = _create_target_docx_run(blank_docx)
    formatting.copy_run_formatting_pptx2docx(
        sub_source, sub_target, UserConfig(experimental_formatting_on=False)
    )
    assert sub_target.font.subscript is not True


# endregion


# endregion

# region pptx2docx PARA tests
# TODO: I guess there's just the one here?

# region TODO: copy_paragraph_formatting_pptx2docx tests
# endregion

# endregion


# region apply_experimental_formatting_from_metadata tests
# endregion

# region helper tests
# TODO: test colormap?
# TODO: test alignment map?
# endregion
