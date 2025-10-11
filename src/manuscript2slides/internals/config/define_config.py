"""TODO: add docstring"""

# ==DOCSTART==
# Purpose: Defines the UserConfig dataclass-- the single source of truth for user-overridable options.
# ==DOCEND==

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from enum import Enum
import os

# TODO: remove later
from manuscript2slides.internals.paths import (
    user_input_dir,
    user_templates_dir,
    user_output_dir,
    user_base_dir,
)


# Which chunking method to use to divide the docx into slides. This enum lists the available choices:
class ChunkType(Enum):
    """Chunk Type Choices"""

    HEADING_NESTED = "heading_nested"
    HEADING_FLAT = "heading_flat"
    PARAGRAPH = "paragraph"
    PAGE = "page"


@dataclass
class UserConfig:
    """All user-configurable settings for manuscript2slides."""

    # Input/Output

    # Input file to process
    input_docx: Optional[str] = (
        None  # Use strings in the dataclass, convert to Path when you need to use them.
    )
    input_pptx: Optional[str] = None

    output_folder: Optional[str] = None  # Desired output directory/folder to save in

    # ==> Templates I/O
    template_pptx: Optional[str] = (
        None  # The pptx file to use as the template for the slide deck
    )
    template_docx: Optional[str] = (
        None  # The docx file to use as the template for the new docx
    )

    # Processing
    chunk_type: ChunkType = (
        ChunkType.PARAGRAPH
    )  # Which chunking method to use to divide the docx into slides.

    experimental_formatting_on: bool = True

    display_comments: bool = True
    comments_sort_by_date: bool = True
    comments_keep_author_and_date: bool = True

    display_footnotes: bool = True
    display_endnotes: bool = True

    # We this way to leave speaker notes completely empty if the user really wants that, it's a valid use case.
    # Documentation and tooltips should make it clear that this means metadata loss for round-trip pipeline data.
    preserve_docx_metadata_in_speaker_notes: bool = True

    # Class methods
    def _resolve_path(self, raw: str) -> Path:
        """Expand ~ and ${VARS}; resolve relative to config_base_dir if present."""
        expanded = os.path.expandvars(raw)
        p = Path(expanded).expanduser()

        if p.is_absolute():
            return p.resolve()

        base = user_base_dir()
        return (base / p).resolve()

    
    def get_template_pptx_path(self) -> Path:
        """Get the docx2pptx template pptx path, with fallback to default."""
        if self.template_pptx:
            return self._resolve_path(self.template_pptx)

        # Default
        base = user_templates_dir()
        return base / "blank_template.pptx"

    def get_template_docx_path(self) -> Path:
        """Get the pptx2docx template docx path with fallback to a default."""
        if self.template_docx:
            return self._resolve_path(self.template_docx)

        # Default
        base = user_templates_dir()
        return base / "docx_template.docx"

    def get_input_docx_file(self) -> Path:
        """Get the docx2pptx input docx file or fall back to a dry run example file."""
        if self.input_docx:
            return self._resolve_path(self.input_docx)

        # Default/Dry Run from user's input folder (copied there from scaffold.py)
        base = user_input_dir()
        return base / "sample_doc.docx"

    def get_output_folder(self) -> Path:
        """Get the docx2pptx pipeline output pptx path, with fallback to default."""
        if self.output_folder:
            return self._resolve_path(self.output_folder)

        # Default
        return user_output_dir()

    def get_input_pptx_file(self) -> Path:
        """Get the pptx2docx input pptx file or fall back to a dry run example pptx."""
        if self.input_pptx:
            return self._resolve_path(self.input_pptx)

        # Default/Dry Run
        base = user_input_dir()
        return base / "sample_slides_output.pptx"

    # Validation
    def validate(self) -> None:
        """
        Validate intrinsic config values (no filesystem access).

        Catches:
            - Someone accidentally passing wrong types
            - Empty strings where None is expected
            - Enum values that shouldn't be possible (though the enum mostly handles this)
        """

        # Validate chunk_type is a valid ChunkType enum member
        if not isinstance(self.chunk_type, ChunkType):
            raise ValueError(
                f"chunk_type must be a ChunkType enum, got {type(self.chunk_type).__name__}. "
                f"Valid values: {[e.value for e in ChunkType]}"
            )

        # Validate boolean fields are actually booleans
        bool_fields = [
            "experimental_formatting_on",
            "display_comments",
            "display_footnotes",
            "display_endnotes",
            "preserve_docx_metadata_in_speaker_notes",
            "comments_sort_by_date",
            "comments_keep_author_and_date",
        ]

        for field_name in bool_fields:
            val = getattr(self, field_name)
            if not isinstance(val, bool):
                raise ValueError(
                    f"{field_name} must be a boolean, got {type(val).__name__}"
                )

        # Path strings should be strings, if provided
        if self.input_docx is not None and not isinstance(self.input_docx, str):
            raise ValueError(
                f"input_docx must be a string, got {type(self.input_docx).__name__}"
            )

        if self.output_folder is not None and not isinstance(self.output_folder, str):
            raise ValueError(
                f"output_folder must be a string, got {type(self.output_folder).__name__}"
            )

        # Can't be empty string
        if self.output_folder == "":
            raise ValueError(
                "output_folder cannot be empty string; use None for default"
            )

    # Methods below validate pipeline requirements, and check:
    #   - Output path that exists but isn't a directory
    #   - Missing input files before pipeline starts
    #   - Missing templates

    def _validate_output_folder(self) -> None:
        """Helper: validate output folder is usable"""
        # Output folder must be creatable (or already exist)
        output_folder = self.get_output_folder()
        if output_folder.exists() and not output_folder.is_dir():
            raise ValueError(
                f"Output path exists but is not a directory: {output_folder}"
            )

    def validate_docx2pptx_pipeline_requirements(self) -> None:
        """
        Validate external dependencies required to run the docx2pptx pipeline.

        Checks external state:
        - Verifies files exist
        - Checks permissions
        - Only runs right before you actually need those resources
        """

        # Always validate - either user file or default sample
        input_path = self.get_input_docx_file()
        if not input_path.exists():
            raise FileNotFoundError(f"Input docx file not found: {input_path}")
        if not input_path.is_file():
            raise ValueError(f"Input docx path is not a file: {input_path}")

        # Always need template
        pptx_template_path = self.get_template_pptx_path()
        if not pptx_template_path.exists():
            raise FileNotFoundError(f"Template not found: {pptx_template_path}")

        self._validate_output_folder()

    def validate_pptx2docx_pipeline_requirements(self) -> None:
        """Validate external dependencies required to run the pptx2docx pipeline."""

        # Always validate - either user file or default sample
        input_path = self.get_input_pptx_file()
        if not input_path.exists():
            raise FileNotFoundError(f"Input pptx not found: {input_path}")
        if not input_path.is_file():
            raise ValueError(f"Not a file: {input_path}")

        docx_template_path = self.get_template_docx_path()
        if not docx_template_path.exists():
            raise FileNotFoundError(f"Template not found: {docx_template_path}")

        self._validate_output_folder()  # Shared check
