# internals/config/define_config.py
"""User configuration dataclass and validation."""

# region imports
from __future__ import annotations

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:
    import tomli as tomllib  # Python 3.10

import tomli_w  # For writing (no stdlib equivalent yet)

import os
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional

from manuscript2slides.internals.paths import (
    user_base_dir,
    user_input_dir,
    user_output_dir,
    user_templates_dir,
)
from manuscript2slides.internals.run_context import get_pipeline_run_id
import logging

log = logging.getLogger("manuscript2slides")
# endregion


# region Enums
# Which chunking method to use to divide the docx into slides. This enum lists the available choices:
class ChunkType(Enum):
    """Chunk Type Choices"""

    HEADING_NESTED = "heading_nested"
    HEADING_FLAT = "heading_flat"
    PARAGRAPH = "paragraph"
    PAGE = "page"


class PipelineDirection(Enum):
    """Pipeline direction choices"""

    DOCX_TO_PPTX = "docx2pptx"
    PPTX_TO_DOCX = "pptx2docx"


# endregion


# region class UserConfig
@dataclass
class UserConfig:
    """All user-configurable settings for manuscript2slides."""

    # endregion

    # region define fields
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

    direction: PipelineDirection = PipelineDirection.DOCX_TO_PPTX

    experimental_formatting_on: bool = True

    display_comments: bool = True
    comments_sort_by_date: bool = True
    comments_keep_author_and_date: bool = True

    display_footnotes: bool = True
    display_endnotes: bool = True

    # We this way to leave speaker notes completely empty if the user really wants that, it's a valid use case.
    # Documentation and tooltips should make it clear that this means metadata loss for round-trip pipeline data.
    preserve_docx_metadata_in_speaker_notes: bool = True

    # endregion

    # region baseline instance methods
    def _resolve_path(self, raw: str) -> Path:
        """Expand ~ and ${VARS}; resolve relative to config_base_dir if present."""
        expanded = os.path.expandvars(raw)
        p = Path(expanded).expanduser()

        if p.is_absolute():
            return p.resolve()

        base = user_base_dir()
        return (base / p).resolve()

    def _make_path_relative(self, path_str: str | None) -> str | None:
        """
        Convert absolute paths under user_base_dir to relative paths for readability.

        Paths outside user_base_dir are kept as absolute paths.
        Uses forward slashes for cross-platform compatibility.
        """
        if path_str is None:
            return None

        abs_path = self._resolve_path(path_str)
        base = user_base_dir()

        try:
            # Try to make it relative to base
            rel_path = abs_path.relative_to(base)
            # Use forward slashes (work on Windows too, cleaner in TOML)
            return str(rel_path).replace("\\", "/")
        except ValueError:
            # Path is outside base dir, keep it absolute with forward slashes
            return str(abs_path).replace("\\", "/")

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

    def get_input_docx_file(self) -> Path | None:
        """Get the docx2pptx input docx file path, or None if not specified."""
        if self.input_docx:
            return self._resolve_path(self.input_docx)

        return None

    def get_output_folder(self) -> Path:
        """Get the docx2pptx pipeline output pptx path, with fallback to default."""
        if self.output_folder:
            return self._resolve_path(self.output_folder)

        # Default
        return user_output_dir()

    def get_input_pptx_file(self) -> Path | None:
        """Get the pptx2docx input pptx file path, or None if not specified."""
        if self.input_pptx:
            return self._resolve_path(self.input_pptx)

        return None  # No more fallback to sample

    # endregion

    # region class methods to populate an instance
    @classmethod
    def with_defaults(cls) -> UserConfig:
        """
        Create a config object in memory with sample files for quick CLI demo.

        Provides sensible defaults for everything so users can run
        `python -m manuscript2slides` and see it work immediately.

        Returns:
            UserConfig: Fully populated config using sample files
        """

        cfg = cls()

        # Point to sample files for demo
        cfg.input_docx = str(user_input_dir() / "sample_doc.docx")

        # All other fields already have defaults from the dataclass
        # (chunk_type, direction, bools, templates, output_folder)
        return cfg

    @classmethod
    def for_demo(cls, direction: PipelineDirection) -> UserConfig:
        """
        Create a config with sample files for GUI demo tab.

        User picks the direction (docx2pptx or pptx2docx), we fill in
        the appropriate sample files and sensible defaults.

        Args:
            direction: Which pipeline direction to demo

        Returns:
            UserConfig: Fully populated config using sample files
        """
        cfg = cls()
        cfg.direction = direction

        # Set the appropriate input file based on direction
        if direction == PipelineDirection.DOCX_TO_PPTX:
            cfg.input_docx = str(user_input_dir() / "sample_doc.docx")
        elif direction == PipelineDirection.PPTX_TO_DOCX:
            cfg.input_pptx = str(user_input_dir() / "sample_slides_output.pptx")

        # All other fields use their dataclass defaults
        # (output_folder, templates, bools, etc.)

        return cfg

    @classmethod
    def from_toml(cls, path: Path) -> UserConfig:
        """
        Load configuration from a TOML file.

        The TOML file should have flat key-value pairs matching the UserConfig field names.

        Example TOML:
            input_docx = "~/my-manuscript.docx"
            chunk_type = "heading_flat"
            direction = "docx2pptx"
            experimental_formatting_on = true

        Args:
            path: Path to the .toml config file

        Returns:
            UserConfig: Populated configuration object

        Raises:
            FileNotFoundError: If config file doesn't exist
            ValueError: If TOML is invalid or contains invalid enum values
        """
        if not path.exists():
            log.error(f"Config file not found: {path}")
            raise FileNotFoundError(f"Config file not found: {path}")

        # Read in the TOML file; raise if there are syntax errors.
        try:
            with open(path, "rb") as f:
                data = tomllib.load(f)
        except tomllib.TOMLDecodeError as e:
            error_msg = f"Invalid TOML syntax in {path}"
            log.error(error_msg)
            raise ValueError(error_msg) from e

        # Warn the user if the data was read-in as empty, but only warn-- keep going.
        if not data:
            log.warning(f"Config toml file is empty: {path}. Using all defaults.")

        # Convert string enum values to actual enums
        # Convert string enum values to actual enums
        if "chunk_type" in data:
            try:
                data["chunk_type"] = ChunkType(data["chunk_type"])
            except ValueError as e:
                error_msg = (
                    f"Invalid chunk_type: '{data['chunk_type']}'. "
                    f"Valid options: {[c.value for c in ChunkType]}"
                )
                log.error(error_msg)
                raise ValueError(error_msg) from e

        if "direction" in data:
            try:
                data["direction"] = PipelineDirection(data["direction"])
            except ValueError as e:
                error_msg = (
                    f"Invalid direction: '{data['direction']}'. "
                    f"Valid options: {[d.value for d in PipelineDirection]}"
                )
                log.error(error_msg)
                raise ValueError(error_msg) from e

        # Create and return UserConfig object from the dict by unpacking all the key-value pairs as kwargs
        return cls(**data)

    # endregion
    # region save_toml
    def save_toml(self, path: Path) -> None:
        """
        Save configuration to a TOML file.

        Args:
            path: Where to save the .toml file
        """
        # Convert to Path if it's a string
        path = Path(path)

        # Check if path is a directory
        if path.exists() and path.is_dir():
            error_msg = f"Cannot save config: path is a directory, not a file: {path}."
            log.error(error_msg)
            raise ValueError(error_msg)

        # Auto-create parent directories if they don't exist
        path.parent.mkdir(parents=True, exist_ok=True)

        # Convert config to dict
        data = {
            "input_docx": self._make_path_relative(self.input_docx),
            "input_pptx": self._make_path_relative(self.input_pptx),
            "output_folder": self._make_path_relative(self.output_folder),
            "template_pptx": self._make_path_relative(self.template_pptx),
            "template_docx": self._make_path_relative(self.template_docx),
            "chunk_type": self.chunk_type.value,
            "direction": self.direction.value,
            "experimental_formatting_on": self.experimental_formatting_on,
            "display_comments": self.display_comments,
            "comments_sort_by_date": self.comments_sort_by_date,
            "comments_keep_author_and_date": self.comments_keep_author_and_date,
            "display_footnotes": self.display_footnotes,
            "display_endnotes": self.display_endnotes,
            "preserve_docx_metadata_in_speaker_notes": self.preserve_docx_metadata_in_speaker_notes,
        }

        # Filter out None values (TOML can't serialize None)
        data = {k: v for k, v in data.items() if v is not None}

        # Write to TOML file
        try:
            with open(path, "wb") as f:
                tomli_w.dump(data, f)
        except PermissionError as e:
            error_msg = f"Permission denied writing to: {path}"
            log.error(error_msg)
            raise PermissionError(error_msg) from e
        except OSError as e:
            error_msg = f"Failed to write config file to {path}"
            log.error(error_msg)
            raise OSError(error_msg) from e

    # endregion

    # region Validation instance methods
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

        # Validate direction is a valid PipelineDirection enum member
        if not isinstance(self.direction, PipelineDirection):
            raise ValueError(
                f"direction must be a PipelineDirection enum, got {type(self.direction).__name__}. "
                f"Valid values: {[e.value for e in PipelineDirection]}"
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

        if self.input_pptx is not None and not isinstance(self.input_pptx, str):
            raise ValueError(
                f"input_pptx must be a string, got {type(self.input_pptx).__name__}"
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

    # =======
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

    # TODO: Add logging to pipeline validation methods (validate_docx2pptx_pipeline_requirements, validate_pptx2docx_pipeline_requirements)
    def validate_docx2pptx_pipeline_requirements(self) -> None:
        """
        Validate external dependencies required to run the docx2pptx pipeline.

        Checks external state:
        - Verifies files exist
        - Checks permissions
        - Only runs right before you actually need those resources
        """

        input_path = self.get_input_docx_file()

        # Check: did user provide an input file at all?
        if input_path is None:
            raise ValueError(
                "No input docx file specified. Please set input_docx before running the pipeline."
            )
        # Check: does the file exist on disk?
        if not input_path.exists():
            raise FileNotFoundError(f"Input docx file not found: {input_path}")
        # Check: is it actually a file (not a directory)?
        if not input_path.is_file():
            raise ValueError(f"Input docx path is not a file: {input_path}")

        # Always need template
        pptx_template_path = self.get_template_pptx_path()
        if not pptx_template_path.exists():
            raise FileNotFoundError(f"Template not found: {pptx_template_path}")

        self._validate_output_folder()

    def validate_pptx2docx_pipeline_requirements(self) -> None:
        """Validate external dependencies required to run the pptx2docx pipeline."""

        # Check: did user provide an input file at all?
        input_path = self.get_input_pptx_file()
        if input_path is None:
            raise ValueError(
                "No input docx file specified. Please set input_pptx before running the pipeline."
            )

        # Check: does the file exist on disk?
        if not input_path.exists():
            raise FileNotFoundError(f"Input pptx not found: {input_path}")

        # Check: is it actually a file (not a directory)?
        if not input_path.is_file():
            raise ValueError(f"Not a file: {input_path}")

        docx_template_path = self.get_template_docx_path()
        if not docx_template_path.exists():
            raise FileNotFoundError(f"Template not found: {docx_template_path}")

        self._validate_output_folder()  # Shared check

    # endregion
