"""User directory structure creation and initialization.
Auto-creates ~/Documents/manuscript2slides/ structure with templates and README

On first run, this creates:
- ~/Documents/manuscript2slides/
  ├── README.md           (explains what each folder is for)
  ├── input/              (optional staging for source files)
  ├── output/             (converted files land here)
  ├── logs/               (manuscript2slides.log lives here)
  ├── configs/            (saved configuration files)
  └── templates/          (pptx_template.pptx, docx_template.docx)

Safe to call repeatedly - won't overwrite existing user files.
"""

import logging
import shutil
from importlib.resources import files
from pathlib import Path

from manuscript2slides.internals.paths import (
    user_base_dir,
    user_configs_dir,
    user_input_dir,
    user_log_dir_path,
    user_manifests_dir,
    user_output_dir,
    user_templates_dir,
)

log = logging.getLogger("manuscript2slides")


# region ensure_user_scaffold
def ensure_user_scaffold() -> None:
    """
    Create folder structure and copy templates on first run.

    Safe to call every time - won't overwrite existing user files.
    Creates:
        - ~/Documents/manuscript2slides/README.md
        - ~/Documents/manuscript2slides/templates/*.pptx
        - ~/Documents/manuscript2slides/templates/*.docx
        - ~/Documents/manuscript2slides/config/sample_config.toml
        - Empty input/output/logs folders
    """

    # Get all directory paths (path functions do NOT create directories)
    base = user_base_dir()
    input_dir = user_input_dir()
    output_dir = user_output_dir()
    log_dir = user_log_dir_path()
    templates = user_templates_dir()
    configs_dir = user_configs_dir()
    manifests_dir = user_manifests_dir()

    # Create all directories explicitly
    base.mkdir(parents=True, exist_ok=True)
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)
    templates.mkdir(parents=True, exist_ok=True)
    configs_dir.mkdir(parents=True, exist_ok=True)
    manifests_dir.mkdir(parents=True, exist_ok=True)

    _create_readme_if_missing(base)

    # Copy template files if missing
    _copy_templates_if_missing(templates)

    # Copy sample docx/pptx files if missing
    _copy_samples_if_missing(input_dir)

    # Copy sample config if missing
    _copy_sample_config_if_missing(configs_dir)

    log.debug(f"User scaffold ready at {base}")


# endregion


# region _get_resource_path
def _get_resource_path(filename: str) -> Path:
    """
    Get path to a packaged resource file used as the source for copying into scaffolded destination subfolders.

    Works in both development and installed scenarios.
    """
    resource = files("manuscript2slides") / "resources" / filename

    # files() returns a Traversable that can be converted to Path
    # In dev: points to src/manuscript2slides/resources/
    # In production: points to site-packages/manuscript2slides/resources/
    return Path(str(resource))


# endregion


# region _create_readme
def _create_readme_if_missing(base_dir: Path) -> Path:
    """Write a friendly README explaining the folder structure if it doesn't already exist."""

    readme_path = base_dir / "README.md"

    if readme_path.exists():
        log.debug("README already exists (not overwriting).")
        return readme_path

    source = _get_resource_path("scaffold_README.md")

    if source.exists():
        readme_text = source.read_text(encoding="utf-8")
        log.debug(f"Found source readme successfully at {source}, copying text.")
    else:
        log.error(f"README template not found: {source}")

        # Fallback: create a minimal README text
        readme_text = "# manuscript2slides\n\nUser folder created automatically.\n"
        log.debug(f"Can't find source README, creating fallback README text instead.")

    readme_path.write_text(readme_text, encoding="utf-8")
    log.info(f"Created new README at {readme_path}")
    return readme_path


# endregion


# region _copy_templates_if_missing
def _copy_templates_if_missing(templates_dir: Path) -> list[Path]:
    """Copy template files from resources/ to user templates folder."""

    templates_to_copy = [
        "pptx_template.pptx",
        "docx_template.docx",
    ]

    paths_processed = []

    for template_name in templates_to_copy:
        source = _get_resource_path(template_name)
        target = templates_dir / template_name

        # Only copy if destination doesn't exist (don't overwrite user customizations)
        if target.exists():
            log.debug(f"Template already exists (not overwriting): {template_name}")
        else:
            if source.exists():
                shutil.copy2(
                    source, target
                )  # copy2 preserves metadata (timestamps, permissions)
                log.info(f"Copied template: {template_name}")
            else:
                log.warning(f"Template not found in resources: {template_name}")

        paths_processed.append(target)

    return paths_processed


# endregion


# region _copy_samples_if_missing
def _copy_samples_if_missing(input_dir: Path) -> list[Path]:
    """Copy sample input files from package resources to user input folder."""

    samples_to_copy = [
        "sample_doc.docx",
        "sample_slides_output.pptx",  # For reverse pipeline testing
    ]

    paths_processed = []

    for sample_name in samples_to_copy:
        source = _get_resource_path(sample_name)
        target = input_dir / sample_name

        # Only copy if destination doesn't exist
        if target.exists():
            log.debug(f"Sample already exists (not overwriting): {sample_name}")

        else:
            if source.exists():
                shutil.copy2(source, target)
                log.info(f"Copied sample: {sample_name}")
            else:
                log.warning(f"Sample not found in resources: {sample_name}")

        paths_processed.append(target)

    return paths_processed


# endregion


# region _copy_sample_config_if_missing
def _copy_sample_config_if_missing(configs_dir: Path) -> Path:
    """Generate sample config with correct absolute paths for this user's system."""

    sample_name = "sample_config.toml"
    target = configs_dir / sample_name

    if target.exists():
        log.debug(f"{target} already exists (not overwriting).")
        return target

    # Generate config content with user-specific absolute paths
    base = user_base_dir()
    sample_docx = base / "input" / "sample_doc.docx"
    output = base / "output"
    template_pptx = base / "templates" / "pptx_template.pptx"

    # Use .as_posix() for cross-platform TOML compatibility
    config_content = f"""# Sample configuration for manuscript2slides
# This file was auto-generated with paths specific to your system

# === Input Files ===
input_docx = "{sample_docx.as_posix()}"
# Uncomment and customize:
# input_docx = "~/Documents/my-manuscript.docx"

# === Output ===
output_folder = "{output.as_posix()}"

# === Templates ===
template_pptx = "{template_pptx.as_posix()}"
# Leave blank to use default template:
# template_pptx = ""

# === Processing Options ===
chunk_type = "paragraph"  # Options: paragraph, page, heading_flat, heading_nested

# === Formatting ===
experimental_formatting_on = true

# === Annotations ===
display_comments = true
comments_sort_by_date = true
comments_keep_author_and_date = true
display_footnotes = true
display_endnotes = true
preserve_docx_metadata_in_speaker_notes = true
"""

    target.write_text(config_content, encoding="utf-8")
    log.info(f"Generated sample config: {sample_name}")
    return target


# endregion
