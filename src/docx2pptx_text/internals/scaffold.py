"""Auto-create user directory structure with templates and documentation.

On first run, this creates:
- ~/Documents/docx2pptx_text/
  ├── README.md           (explains what each folder is for)
  ├── input/              (optional staging for source files)
  ├── output/             (converted files land here)
  ├── logs/               (docx2pptx_text.log lives here)
  └── templates/          (blank_template.pptx, docx_template.docx)

Safe to call repeatedly - won't overwrite existing user files.
"""

# ==DOCSTART==
# Purpose: Auto-create ~/Documents/docx2pptx_text/ structure with templates and README
# ==DOCEND==

import logging
import shutil
from pathlib import Path

from docx2pptx_text.internals.paths import (
    user_base_dir,
    user_input_dir,
    user_output_dir,
    user_log_dir_path,
    user_templates_dir,
)
from docx2pptx_text.internals.constants import RESOURCES_DIR

log = logging.getLogger("docx2pptx_text")


def ensure_user_scaffold() -> None:
    """
    Create folder structure and copy templates on first run.

    Safe to call every time - won't overwrite existing user files.
    Creates:
        - ~/Documents/docx2pptx_text/README.md
        - ~/Documents/docx2pptx_text/templates/*.pptx
        - ~/Documents/docx2pptx_text/templates/*.docx
        - Empty input/output/logs folders
    """

    base = user_base_dir()

    # Ensure all the directories exist (paths.py functions do the mkdir stuff)
    input_dir = user_input_dir()
    user_output_dir()
    user_log_dir_path()
    templates = user_templates_dir()

    readme_path = base / "README.md"
    if not readme_path.exists():
        _create_readme(readme_path)
        log.info(f"Created new README at {readme_path}")

    # Copy template files if missing
    _copy_templates_if_missing(templates)

    # Copy sample files if missing
    _copy_samples_if_missing(input_dir)

    log.debug(f"User scaffold ready at {base}")


def _create_readme(path: Path) -> None:
    """Write a friendly README explaining the folder structure."""

    # TODO: When packaging, use importlib.resources instead of ROOT_DIR
    source = RESOURCES_DIR / "scaffold_README.md"

    if source.exists():
        readme_text = source.read_text(encoding="utf-8")
        path.write_text(readme_text, encoding="utf-8")
    else:
        log.error(f"README template not found: {source}")

        # Fallback: create a minimal README
        path.write_text(
            "# docx2pptx_text\n\nUser folder created automatically.\n", encoding="utf-8"
        )


def _copy_templates_if_missing(templates_dir: Path) -> None:
    """Copy template files from resources/ to user templates folder."""
    # TODO: When packaging, use importlib.resources instead of ROOT_DIR
    # See: https://docs.python.org/3/library/importlib.resources.html

    # Source: your project's resources folder
    source_dir = RESOURCES_DIR

    templates_to_copy = [
        "blank_template.pptx",
        "docx_template.docx",
    ]

    for template_name in templates_to_copy:
        source = source_dir / template_name
        target = templates_dir / template_name

        # Only copy if destination doesn't exist (don't overwrite user customizations)
        if not target.exists():
            if source.exists():
                shutil.copy2(
                    source, target
                )  # copy2 preserves metadata (timestamps, permissions)
                log.info(f"Copied template: {template_name}")
            else:
                log.warning(f"Template not found in resources: {template_name}")
        else:
            log.debug(f"Template already exists (not overwriting): {template_name}")


def _copy_samples_if_missing(input_dir: Path) -> None:
    """Copy sample input files from package resources to user input folder."""

    samples_to_copy = [
        "sample_doc.docx",
        "sample_slides_output.pptx",  # For reverse pipeline testing
    ]

    for sample_name in samples_to_copy:
        # TODO: When packaging, use importlib.resources instead of RESOURCES_DIR
        source = RESOURCES_DIR / sample_name
        dest = input_dir / sample_name

        # Only copy if destination doesn't exist
        if not dest.exists():
            if source.exists():
                shutil.copy2(source, dest)
                log.info(f"Copied sample: {sample_name}")
            else:
                log.warning(f"Sample not found in resources: {sample_name}")
        else:
            log.debug(f"Sample already exists (not overwriting): {sample_name}")


# =====================

# TODO: When packaging, use importlib.resources instead of RESOURCES_DIR
from importlib.resources import files  # Python 3.9+


def _get_resource_path(filename: str) -> Path:
    """Get path to a packaged resource file."""
    # In development: points to your resources/ folder
    # When installed: points to site-packages/docx2pptx_text/resources/
    return files("docx2pptx_text").joinpath(
        "resources", filename
    )  # pyright: ignore # this gets mad about traversal object can't be path or something
