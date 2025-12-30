# internals/config/define_config.py
"""User configuration dataclass and validation."""

# region imports
from __future__ import annotations

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:
    # Python 3.10
    import tomli as tomllib  # type: ignore[no-redef]

import logging
from dataclasses import dataclass, fields
from enum import Enum
from pathlib import Path
from typing import Any, Optional

import tomli_w  # For writing (no stdlib equivalent yet)

from manuscript2slides.internals.paths import (
    get_default_docx_template_path,
    get_default_pptx_template_path,
    resolve_path,
    user_input_dir,
    user_output_dir,
)

# endregion

log = logging.getLogger("manuscript2slides")


# region Enums
# Which chunking method to use to divide the docx into slides. This enum lists the available choices:
class ChunkType(Enum):
    """Chunk Type Choices"""

    HEADING_NESTED = "heading_nested"
    HEADING_FLAT = "heading_flat"
    PARAGRAPH = "paragraph"
    PAGE = "page"

    @classmethod
    def from_string(cls, value: str) -> "ChunkType":
        """Convert string to ChunkType, with support for aliases."""
        value = value.lower().strip()

        # Alias mapping
        aliases = {
            "heading": cls.HEADING_FLAT,
            # Think hard before adding more here. You'll also have to add them to the CLI argparser.
        }

        # Check aliases first
        if value in aliases:
            return aliases[value]

        for member in cls:
            if member.value == value:
                return member

        valid_values = [m.value for m in cls] + list(aliases.keys())
        raise ValueError(
            f"'{value}' is not a valid ChunkType. Valid options: {', '.join(valid_values)}"
        )


class PipelineDirection(Enum):
    """Pipeline direction choices"""

    DOCX_TO_PPTX = "docx2pptx"
    PPTX_TO_DOCX = "pptx2docx"


# endregion


# region class UserConfig
@dataclass
class UserConfig:
    """All user-configurable settings for manuscript2slides."""

    # region class fields

    # region Input/Output
    # Input file to process
    input_docx: Optional[Path] = None
    input_pptx: Optional[Path] = None

    output_folder: Optional[Path] = None  # Desired output directory/folder to save in

    # Templates for output file
    template_pptx: Optional[Path] = (
        None  # The pptx file to use as the template for the slide deck
    )
    template_docx: Optional[Path] = (
        None  # The docx file to use as the template for the new docx
    )
    range_start: Optional[int] = None
    range_end: Optional[int] = None
    # endregion

    # region Processing options
    chunk_type: ChunkType = (
        ChunkType.PARAGRAPH
    )  # Which chunking method to use to divide the docx into slides.

    experimental_formatting_on: bool = True

    display_comments: bool = False
    comments_sort_by_date: bool = True
    comments_keep_author_and_date: bool = True

    display_footnotes: bool = False
    display_endnotes: bool = False

    # We this way to leave speaker notes completely empty if the user really wants that, it's a valid use case.
    # Documentation and tooltips should make it clear that this means metadata loss for round-trip pipeline data.
    preserve_docx_metadata_in_speaker_notes: bool = False

    # endregion

    # endregion

    # region post_init
    def __post_init__(self) -> None:
        """Convert string inputs of path fields into Path objects."""
        if self.input_docx is not None:
            self.input_docx = Path(self.input_docx)
        if self.input_pptx is not None:
            self.input_pptx = Path(self.input_pptx)
        if self.output_folder is not None:
            self.output_folder = Path(self.output_folder)
        if self.template_pptx is not None:
            self.template_pptx = Path(self.template_pptx)
        if self.template_docx is not None:
            self.template_docx = Path(self.template_docx)

    # endregion

    # region class methods (populate a new instance)

    # region with_defaults
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
        cfg.input_docx = user_input_dir() / "sample_doc.docx"

        # All other fields already have defaults from the dataclass
        # (chunk_type, direction, bools, templates, output_folder)
        return cfg

    # endregion

    # region for_demo
    @classmethod
    def for_demo(cls, requested_direction: PipelineDirection) -> UserConfig:
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

        # Set the appropriate input file based on direction
        if requested_direction == PipelineDirection.DOCX_TO_PPTX:
            cfg.input_docx = user_input_dir() / "sample_doc.docx"
        elif requested_direction == PipelineDirection.PPTX_TO_DOCX:
            cfg.input_pptx = user_input_dir() / "sample_slides_output.pptx"

        # All other fields use their dataclass defaults
        # (output_folder, templates, bools, etc.)

        return cfg

    # endregion

    # region from_toml
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
            error_msg = f"Config file not found: {path}"
            log.error(error_msg)
            raise FileNotFoundError(error_msg)
        if path.is_dir():
            error_msg = f"This is a directory (folder), not a toml file: {path}"
            log.error(error_msg)
            raise ValueError(error_msg)

        # Read in the TOML file; raise if there are syntax errors.
        try:
            with open(path, "rb") as f:
                data = tomllib.load(f)
        except tomllib.TOMLDecodeError as e:
            error_msg = f"Invalid TOML syntax in {path}. Check for missing or mismatched quote marks."
            log.error(error_msg)
            raise ValueError(error_msg) from e
        except PermissionError as e:
            error_msg = f"We hit a permission error when trying to access {path}"
            log.error(error_msg)
            raise ValueError(error_msg) from e

        # Warn the user if the data was read-in as empty, but only warn-- keep going.
        if not data:
            log.warning(
                f"Config toml file loaded as empty, so no UserConfig fields were set from: {path}."
            )
            log.warning(
                f"If you're unhappy this was only logged as a warning and did not stop pipeline execution, please send us that UX feedback via a GitHub Issue."
            )

        # Check for invalid keys
        # Filter out any unexpected fields and warn/error
        valid_fields = {f.name for f in fields(cls)}
        unexpected = set(data.keys()) - valid_fields

        if unexpected:
            log.warning(
                f"Ignoring unexpected fields in TOML config: {', '.join(sorted(unexpected))}. "
                f"Check for typos. Valid fields: {', '.join(sorted(valid_fields))}"
            )

            # Remove unexpected fields
            log.warning("Filtering out unexpected fields before continuing.")
            data = {k: v for k, v in data.items() if k in valid_fields}

        # Convert string enum values to actual enums
        if "chunk_type" in data:
            try:
                data["chunk_type"] = ChunkType.from_string(data["chunk_type"])
            except ValueError as e:
                error_msg = (
                    f"Invalid chunk_type: '{data['chunk_type']}'. "
                    f"Valid options: {[c.value for c in ChunkType]}"
                )
                log.error(error_msg)
                raise ValueError(error_msg) from e

        # Create and return UserConfig object from the dict by unpacking all the key-value pairs as kwargs
        return cls(**data)

    # endregion

    # endregion

    # region instance getters/setters/helpers

    # region direction property
    @property
    def direction(self) -> PipelineDirection:
        """Direction inferred from which input file is set."""
        if self.input_docx and self.input_pptx:
            log.error(
                "We couldn't set the direction of the pipeline because both input_docx and input_pptx were provided. Only 1 input can be specified per run."
            )
            raise ValueError("Cannot determine direction: too many inputs provided")
        elif self.input_docx:
            return PipelineDirection.DOCX_TO_PPTX
        elif self.input_pptx:
            return PipelineDirection.PPTX_TO_DOCX
        else:
            log.error(
                "We couldn't set the direction of the pipeline because there was no input_docx or input_pptx path provided. You must provide at least one."
            )
            raise ValueError("Cannot determine direction: no input file specified")

    # endregion

    # region get real Path objects from stored cfg str values
    def get_template_pptx_path(self) -> Path:
        """Get the docx2pptx template pptx path, with fallback to default."""
        if self.template_pptx:
            return resolve_path(self.template_pptx)

        # Default
        return get_default_pptx_template_path()

    def get_template_docx_path(self) -> Path:
        """Get the pptx2docx template docx path with fallback to a default."""
        if self.template_docx:
            return resolve_path(self.template_docx)

        # Default
        return get_default_docx_template_path()

    def get_input_docx_file(self) -> Path | None:
        """Get the docx2pptx input docx file path, or None if not specified."""
        if self.input_docx:
            return resolve_path(self.input_docx)

        return None

    def get_output_folder(self) -> Path:
        """Get the docx2pptx pipeline output pptx path, with fallback to default."""
        if self.output_folder:
            return resolve_path(self.output_folder)

        # Default
        return user_output_dir()

    def get_input_pptx_file(self) -> Path | None:
        """Get the pptx2docx input pptx file path, or None if not specified."""
        if self.input_pptx:
            return resolve_path(self.input_pptx)

        return None

    def get_input_file(self) -> Path | None:
        """Get the input file path as a Path."""
        if self.direction == PipelineDirection.DOCX_TO_PPTX:
            path = self.get_input_docx_file()
        elif self.direction == PipelineDirection.PPTX_TO_DOCX:
            path = self.get_input_pptx_file()
        else:
            path = None
        return path if path else None

    # endregion

    # region enable_all_options
    def enable_all_options(self) -> UserConfig:
        """
        Enable all processing options by mutating an existing config.
        Use for demos, testing, and preserving as much as possible during
        roundtrip pipeline runs. Returns the mutated object.
        """
        log.info("Enabling all processing bool options on existing config.")

        # Defaults should already set these to True, but just in case.
        self.experimental_formatting_on = True
        self.comments_keep_author_and_date = True
        self.comments_sort_by_date = True

        # We leave these empty by default to have speaker notes empty,
        # but for demo and testing cases we want them enabled.
        self.preserve_docx_metadata_in_speaker_notes = True
        self.display_comments = True
        self.display_endnotes = True
        self.display_footnotes = True

        return self

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
        data: dict[str, Any] = self.config_to_dict()

        # Filter out None values (TOML can't serialize None)
        data: dict[str, Any] = {k: v for k, v in data.items() if v is not None}  # type: ignore[no-redef]
        log.debug(f"Data to be written to toml file is: \n{data}")

        # Write to TOML file
        try:
            log.info(f"Attempting to save to {path}")
            with open(path, "wb") as f:
                tomli_w.dump(data, f)
            log.info(f"Saved toml config file at {path}")
        except PermissionError as e:
            error_msg = f"Permission denied writing to: {path}"
            log.error(error_msg)
            raise PermissionError(error_msg) from e
        except OSError as e:
            error_msg = f"Failed to write config file to {path}"
            log.error(error_msg)
            raise OSError(error_msg) from e

    # endregion

    # region config_to_dict
    def config_to_dict(self) -> dict[str, Any]:
        """Convert config to a TOML-serializable dict.
        Path separators are normalized cross-platform
        portability and to avoid TOML escape sequence issues.
        """

        data: dict[str, Any] = {
            "input_docx": (
                self.input_docx.as_posix() if self.input_docx else None
            ),  # Convert Path to string with .as_posix() - gives you forward slashes (cross-platform)
            "input_pptx": self.input_pptx.as_posix() if self.input_pptx else None,
            "output_folder": (
                self.output_folder.as_posix() if self.output_folder else None
            ),
            "template_pptx": (
                self.template_pptx.as_posix() if self.template_pptx else None
            ),
            "template_docx": (
                self.template_docx.as_posix() if self.template_docx else None
            ),
            "range_start": self.range_start,
            "range_end": self.range_end,
            "chunk_type": self.chunk_type.value,
            "experimental_formatting_on": self.experimental_formatting_on,
            "display_comments": self.display_comments,
            "comments_sort_by_date": self.comments_sort_by_date,
            "comments_keep_author_and_date": self.comments_keep_author_and_date,
            "display_footnotes": self.display_footnotes,
            "display_endnotes": self.display_endnotes,
            "preserve_docx_metadata_in_speaker_notes": self.preserve_docx_metadata_in_speaker_notes,
        }

        log.debug(f"Data to be written is: \n{data}")

        return data

    # endregion

    # endregion

    # region instance validation methods
    def pre_run_check(self) -> None:
        """
        Validate everything needed for a pipeline run.
        Combines intrinsic and external validation in one place.
        """
        # Intrinsic validation: This is likely to have already been done by a caller for UX reasons, but we repeat it here
        # out of caution and for the sake of correctness.
        self.validate()

        # External validation based on pipeline direction
        if self.direction == PipelineDirection.DOCX_TO_PPTX:
            self.validate_docx2pptx_pipeline_requirements()
        elif self.direction == PipelineDirection.PPTX_TO_DOCX:
            self.validate_pptx2docx_pipeline_requirements()
        else:
            raise ValueError(f"Unknown pipeline direction: {self.direction}")

    def validate(self) -> None:
        """
        Validate intrinsic config values (no filesystem access).

        Catches:
            - Someone accidentally passing wrong types
            - Empty strings where None is expected
            - Enum values that shouldn't be possible (though the enum mostly handles this)
        """

        # Can't have both inputs set simultaneously
        if self.input_docx and self.input_pptx:
            log.error("Cannot specify both input_docx and input_pptx")
            raise ValueError(
                "Cannot specify both input_docx and input_pptx. "
                "Please specify only one input file."
            )

        # Must have at least one input set:
        if not self.input_docx and not self.input_pptx:
            log.error("No input file specified")
            raise ValueError(
                "No input file provided: Must specify either input_docx or input_pptx."
            )

        # Warn if the user accidentally passed the same file type for input + template (template will be ignored)
        # TODO, post-release: Assess if this should raise/fail the pipeline rather than warn and just ignore the input.
        if self.input_docx and self.template_docx:
            log.warning(
                "You provided a template_docx, but you passed an input_docx for the pipeline, which means the docx2pptx pipeline will run, and your template input will be ignored. (We will use the default template_pptx.)\n"
                "If you're reading this in the log after a bunch of frustration of trying to get your template to work, we probably should have errored-out and failed rather than just logging a warning. Please let us know on the github if that's the behavior you would've preferred.",
            )
        if self.input_pptx and self.template_pptx:
            log.warning(
                "You provided a template_pptx, but you passed an input_pptx for the pipeline, which means the pptx2docx pipeline will run, and your template input will be ignored. (We will use the default template_docx)\n"
                "If you're reading this in the log after a bunch of frustration of trying to get your template to work, we probably should have errored-out and failed rather than just logging a warning. Please let us know on the github if that's the behavior you would've preferred.",
            )

        # Validate chunk_type is a valid ChunkType enum member
        assert isinstance(self.chunk_type, ChunkType), (
            f"chunk_type must be ChunkType, got {type(self.chunk_type).__name__}"
        )
        if not isinstance(self.chunk_type, ChunkType):  # type: ignore[unreachable]
            log.error("Invalid value in chunk_type; must be enum.")
            raise ValueError(
                f"chunk_type must be a ChunkType enum, got {type(self.chunk_type).__name__}. "
                f"Valid values: {[e.value for e in ChunkType]}"
            )

        # Validate direction is a valid PipelineDirection enum member
        if not isinstance(self.direction, PipelineDirection):  # type: ignore[unreachable]
            log.error("Invalid value in direction; must be enum.")
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
                log.error(f"{field_name} must be a boolean, got {type(val).__name__}")
                raise ValueError(
                    f"{field_name} must be a boolean, got {type(val).__name__}"
                )

        if self.range_start is not None:
            if not isinstance(self.range_start, int):  # type: ignore[unreachable]
                log.error(
                    f"range_start must be an integer, got {type(self.range_start).__name__}"
                )
                raise ValueError(
                    f"range_start must be an integer, got {type(self.range_start).__name__}"
                )
            if self.range_start < 1:
                log.error(f"range_start must be >= 1, got {self.range_start}")
                raise ValueError(f"range_start must be >= 1, got {self.range_start}")

        if self.range_end is not None:
            if not isinstance(self.range_end, int):  # type: ignore[unreachable]
                log.error(
                    f"range_end must be an integer, got {type(self.range_end).__name__}"
                )
                raise ValueError(
                    f"range_end must be an integer, got {type(self.range_end).__name__}"
                )
            if self.range_end < 1:
                log.error(f"range_end must be >= 1, got {self.range_end}")
                raise ValueError(f"range_end must be >= 1, got {self.range_end}")

        # Validate start + end range logic
        if self.range_start is not None and self.range_end is not None:
            if self.range_start > self.range_end:
                error_msg = f"range_start ({self.range_start}) cannot be greater than range_end ({self.range_end})"
                log.error(error_msg)
                raise ValueError(error_msg)

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

    def validate_docx2pptx_pipeline_requirements(self) -> None:
        """
        Validate external dependencies required to run the docx2pptx pipeline.

        Checks external state:
        - Verifies files & directories exist
        - Verifies expected file/folder type
        - Only runs right before you actually need those resources
        """

        input_path = self.get_input_docx_file()

        # Check: did user provide an input file at all?
        if input_path is None:
            log.error("No input docx file specified.")
            raise ValueError(
                "No input docx file specified. Please set input_docx before running the pipeline."
            )
        # Check: does the file exist on disk?
        if not input_path.exists():
            error_msg = f"Input docx file not found: {input_path}"
            log.error(error_msg)
            raise FileNotFoundError(error_msg)
        # Check: is it actually a file (not a directory)?
        if not input_path.is_file():
            error_msg = f"Input docx path is not a file: {input_path}"
            log.error(error_msg)
            raise ValueError(error_msg)

        # Always need template
        pptx_template_path = self.get_template_pptx_path()
        if not pptx_template_path.exists():
            error_msg = f"Template not found: {pptx_template_path}"
            log.error(error_msg)
            raise FileNotFoundError(error_msg)
        elif pptx_template_path.is_dir():
            error_msg = f"Template must be a file, not a folder: {pptx_template_path}"
            log.error(error_msg)
            raise ValueError(error_msg)

        self._validate_output_folder()

    def validate_pptx2docx_pipeline_requirements(self) -> None:
        """Validate external dependencies required to run the pptx2docx pipeline."""

        # Check: did user provide an input file at all?
        input_path = self.get_input_pptx_file()
        if input_path is None:
            log.error("No input pptx file specified.")
            raise ValueError(
                "No input pptx file specified. Please set input_pptx before running the pipeline."
            )

        # Check: does the file exist on disk?
        if not input_path.exists():
            error_msg = f"Input pptx file not found: {input_path}"
            log.error(error_msg)
            raise FileNotFoundError(error_msg)

        # Check: is it actually a file (not a directory)?
        if not input_path.is_file():
            error_msg = f"Input pptx is not a file: {input_path}"
            log.error(error_msg)
            raise ValueError(error_msg)

        docx_template_path = self.get_template_docx_path()
        if not docx_template_path.exists():
            error_msg = f"Template not found: {docx_template_path}"
            log.error(error_msg)
            raise FileNotFoundError(error_msg)
        elif docx_template_path.is_dir():
            error_msg = f"Template must be a file, not a folder: {docx_template_path}"
            log.error(error_msg)
            raise ValueError(error_msg)

        self._validate_output_folder()  # Shared check

    # endregion


# endregion
