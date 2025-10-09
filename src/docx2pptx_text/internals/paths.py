"""Cross-platform path resolution for user directories.

Uses platformdirs to find OS-appropriate locations for:
- Logs (where docx2pptx.log lives)
- Output (default save location for converted files)
- Input (optional staging area for source files)
- Templates (custom pptx/docx templates)
"""

# ==DOCSTART==
# Purpose: Utilities for resolving environment-specific paths for logs, output, templates, and input staging.
# ==DOCEND==

from pathlib import Path
from platformdirs import (
    user_documents_dir,
)  # Gives us the "right" place for files on each OS

PACKAGE_NAME = "docx2pptx"


def user_base_dir() -> Path:
    """
    Base directory for all docx2pptx_text user files.

    Returns:
        Path to ~/Documents/docx2pptx/ (or OS equivalent)

    Examples:
        Windows: C:/Users/YourName/Documents/docx2pptx/
        macOS: /Users/YourName/Documents/docx2pptx/
        Linux: /home/yourname/Documents/docx2pptx/
    """
    base = Path(user_documents_dir()) / PACKAGE_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


def user_log_dir_path() -> Path:
    """
    Directory for log files.

    Returns:
        Path to ~/Documents/docx2pptx/logs/
    """
    log_dir = user_base_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def user_output_dir() -> Path:
    """
    Default output directory for converted files.

    Returns:
        Path to ~/Documents/docx2pptx/output/
    """
    output_dir = user_base_dir() / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def user_input_dir() -> Path:
    """
    Optional staging directory for input files.

    Returns:
        Path to ~/Documents/docx2pptx/input/
    """
    input_dir = user_base_dir() / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    return input_dir


def user_templates_dir() -> Path:
    """
    Directory for custom template files.

    Returns:
        Path to ~/Documents/docx2pptx/templates/
    """
    templates_dir = user_base_dir() / "templates"
    templates_dir.mkdir(parents=True, exist_ok=True)
    return templates_dir
