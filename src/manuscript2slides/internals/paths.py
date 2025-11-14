"""Cross-platform path resolution for user directories.

Uses platformdirs to find OS-appropriate locations for:
- Logs (where docx2pptx.log lives)
- Output (default save location for converted files)
- Input (optional staging area for source files)
- Templates (custom pptx/docx templates)
"""

from pathlib import Path
from platformdirs import (
    user_documents_dir,
)  # Gives us the "right" place for files on each OS
import os

PACKAGE_NAME = "manuscript2slides"


# region user_base_dir
def user_base_dir() -> Path:
    """
    Base directory for all manuscript2slides user files.

    Returns:
        Path to ~/Documents/manuscript2slides/ (or OS equivalent)

    Examples:
        Windows: C:/Users/YourName/Documents/manuscript2slides/
        macOS: /Users/YourName/Documents/manuscript2slides/
        Linux: /home/yourname/Documents/manuscript2slides/
    """
    base = Path(user_documents_dir()) / PACKAGE_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


# endregion


# region user_log_dir_path
def user_log_dir_path() -> Path:
    """
    Directory for log files.

    Returns:
        Path to ~/Documents/manuscript2slides/logs/
    """
    log_dir = user_base_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


# endregion


# region user_output_dir
def user_output_dir() -> Path:
    """
    Default output directory for converted files.

    Returns:
        Path to ~/Documents/manuscript2slides/output/
    """
    output_dir = user_base_dir() / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


# endregion


# region user_input_dir
def user_input_dir() -> Path:
    """
    Optional staging directory for input files.

    Returns:
        Path to ~/Documents/manuscript2slides/input/
    """
    input_dir = user_base_dir() / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    return input_dir


# endregion


# region user_templates_dir
def user_templates_dir() -> Path:
    """
    Directory for custom template files.

    Returns:
        Path to ~/Documents/manuscript2slides/templates/
    """
    templates_dir = user_base_dir() / "templates"
    templates_dir.mkdir(parents=True, exist_ok=True)
    return templates_dir


# endregion


# region user_configs_dir
def user_configs_dir() -> Path:
    """
    Directory for saved configuration files.

    Returns:
        Path to ~/Documents/manuscript2slides/configs/
    """
    configs_dir = user_base_dir() / "configs"
    configs_dir.mkdir(parents=True, exist_ok=True)
    return configs_dir


# endregion


# region user_manifests_dir
def user_manifests_dir() -> Path:
    """
    Directory for saved manifest files.

    Returns:
        Path to ~/Documents/manuscript2slides/manifests/
    """
    configs_dir = user_base_dir() / "manifests"
    configs_dir.mkdir(parents=True, exist_ok=True)
    return configs_dir


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
def resolve_path(raw: str) -> Path:
    """
    Expand ~ and ${VARS}; resolve to absolute path.

    Relative paths resolve relative to current working directory.
    """
    expanded = os.path.expandvars(raw)
    return Path(expanded).expanduser().resolve()


# endregion


# region normalize_path
def normalize_path(path_str: str | None) -> str | None:
    """
    Normalize path separators to forward slashes for cross-platform compatibility.

    Forward slashes work on all platforms (Windows, Mac, Linux) and avoid
    TOML escape sequence issues with backslashes.

    Args:
        path_str: Path string (may contain backslashes on Windows)

    Returns:
        Path with forward slashes, or None if input was None
    """
    return path_str.replace("\\", "/") if path_str else None


# endregion
