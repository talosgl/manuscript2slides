"""User directory structure creation and initialization.
Auto-creates ~/Documents/manuscript2slides/ structure with templates and README

On first run, this creates:
- ~/Documents/manuscript2slides/
  ├── README.md           (explains what each folder is for)
  ├── input/              (optional staging for source files)
  ├── output/             (converted files land here)
  ├── logs/               (manuscript2slides.log lives here)
  └── templates/          (blank_template.pptx, docx_template.docx)

Safe to call repeatedly - won't overwrite existing user files.
"""

import logging
import shutil
from pathlib import Path

from manuscript2slides.internals.paths import (
    user_base_dir,
    user_input_dir,
    user_output_dir,
    user_log_dir_path,
    user_templates_dir,
)
from manuscript2slides.internals.constants import RESOURCES_DIR

log = logging.getLogger("manuscript2slides")


def ensure_user_scaffold() -> None:
    """
    Create folder structure and copy templates on first run.

    Safe to call every time - won't overwrite existing user files.
    Creates:
        - ~/Documents/manuscript2slides/README.md
        - ~/Documents/manuscript2slides/templates/*.pptx
        - ~/Documents/manuscript2slides/templates/*.docx
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


# from importlib.resources import files # TODO: uncomment when packaging
def _get_resource_path(filename: str) -> Path:
    """
    Get path to a packaged resource file used as the source for copying into scaffolded destination subfolders.

    Currently uses direct filesystem access for development (./src/manuscript2slides/resources)
    When installed will to site-packages/manuscript2slides/resources/
    """
    return RESOURCES_DIR / filename

    # TODO: When packaging, switch to the below:

    # Convert Traversable to Path
    # resource = files("manuscript2slides").joinpath("resources", filename)
    # return Path(str(resource))


def _create_readme(path: Path) -> None:
    """Write a friendly README explaining the folder structure."""

    source = _get_resource_path("scaffold_README.md")

    if source.exists():
        readme_text = source.read_text(encoding="utf-8")
        path.write_text(readme_text, encoding="utf-8")
    else:
        log.error(f"README template not found: {source}")

        # Fallback: create a minimal README
        path.write_text(
            "# manuscript2slides\n\nUser folder created automatically.\n",
            encoding="utf-8",
        )


def _copy_templates_if_missing(templates_dir: Path) -> None:
    """Copy template files from resources/ to user templates folder."""

    templates_to_copy = [
        "blank_template.pptx",
        "docx_template.docx",
    ]

    for template_name in templates_to_copy:
        source = _get_resource_path(template_name)
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
        source = _get_resource_path(sample_name)
        target = input_dir / sample_name

        # Only copy if destination doesn't exist
        if not target.exists():
            if source.exists():
                shutil.copy2(source, target)
                log.info(f"Copied sample: {sample_name}")
            else:
                log.warning(f"Sample not found in resources: {sample_name}")
        else:
            log.debug(f"Sample already exists (not overwriting): {sample_name}")
