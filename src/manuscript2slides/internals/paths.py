"""Cross-platform path resolution for user directories.

Uses platformdirs to find OS-appropriate locations for:
- Logs (where docx2pptx.log lives)
- Output (default save location for converted files)
- Input (optional staging area for source files)
- Templates (custom pptx/docx templates)
"""

import os
from pathlib import Path

from platformdirs import (
    user_documents_dir,
)  # Gives us the "right" place for files on each OS

PACKAGE_NAME = "manuscript2slides"


# region user_base_dir
def user_base_dir() -> Path:
    """
    Base directory for all manuscript2slides user files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Can be overridden with MANUSCRIPT2SLIDES_BASE_DIR environment variable.

    Returns:
        Path to ~/Documents/manuscript2slides/ (or OS equivalent)

    Examples:
        Windows: C:/Users/YourName/Documents/manuscript2slides/
        macOS: /Users/YourName/Documents/manuscript2slides/
        Linux: /home/yourname/Documents/manuscript2slides/

        With env var override:
        MANUSCRIPT2SLIDES_BASE_DIR=/tmp/test python -m manuscript2slides
    """
    override = os.getenv("MANUSCRIPT2SLIDES_BASE_DIR")
    if override:
        return Path(override)
    return Path(user_documents_dir()) / PACKAGE_NAME


# endregion


# region user_log_dir_path
def user_log_dir_path() -> Path:
    """
    Directory for log files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Returns:
        Path to ~/Documents/manuscript2slides/logs/
    """
    return user_base_dir() / "logs"


# endregion


# region user_output_dir
def user_output_dir() -> Path:
    """
    Default output directory for converted files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Returns:
        Path to ~/Documents/manuscript2slides/output/
    """
    return user_base_dir() / "output"


# endregion


# region user_input_dir
def user_input_dir() -> Path:
    """
    Optional staging directory for input files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Returns:
        Path to ~/Documents/manuscript2slides/input/
    """
    return user_base_dir() / "input"


# endregion


# region user_templates_dir
def user_templates_dir() -> Path:
    """
    Directory for custom template files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Returns:
        Path to ~/Documents/manuscript2slides/templates/
    """
    return user_base_dir() / "templates"


# endregion


# region user_configs_dir
def user_configs_dir() -> Path:
    """
    Directory for saved configuration files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Returns:
        Path to ~/Documents/manuscript2slides/configs/
    """
    return user_base_dir() / "configs"


# endregion


# region user_manifests_dir
def user_manifests_dir() -> Path:
    """
    Directory for saved manifest files.

    Returns path - does NOT create it. Call ensure_user_scaffold() to create.

    Returns:
        Path to ~/Documents/manuscript2slides/manifests/
    """
    return user_base_dir() / "manifests"


# endregion


# region get_default_docx_template_path
def get_default_docx_template_path() -> Path:
    """The default path used for the docx template in the ppt2docx pipeline if none is provided by the user.
    This file is created by scaffold.py's _copy_templates_if_missing() function if it doesn't exist already.
    """
    base = user_templates_dir()
    return base / "docx_template.docx"


# endregion


# region get_default_pptx_template_path
def get_default_pptx_template_path() -> Path:
    """The default path used for the pptx template in the docx2pptx pipeline if none is provided by the user.
    This file is created by scaffold.py's _copy_templates_if_missing() function if it doesn't exist already.
    """
    base = user_templates_dir()
    return base / "pptx_template.pptx"


# endregion


# region resolve_path
def resolve_path(raw: str | Path) -> Path:
    """
    Expand ~ and ${VARS}; resolve to absolute path.

    Relative paths resolve relative to current working directory.
    """
    if isinstance(raw, Path):
        return raw.expanduser().resolve()
    expanded = os.path.expandvars(raw)
    return Path(expanded).expanduser().resolve()


# endregion
