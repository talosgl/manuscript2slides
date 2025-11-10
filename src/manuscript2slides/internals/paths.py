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

PACKAGE_NAME = "manuscript2slides"


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


def user_log_dir_path() -> Path:
    """
    Directory for log files.

    Returns:
        Path to ~/Documents/manuscript2slides/logs/
    """
    log_dir = user_base_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def user_output_dir() -> Path:
    """
    Default output directory for converted files.

    Returns:
        Path to ~/Documents/manuscript2slides/output/
    """
    output_dir = user_base_dir() / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def user_input_dir() -> Path:
    """
    Optional staging directory for input files.

    Returns:
        Path to ~/Documents/manuscript2slides/input/
    """
    input_dir = user_base_dir() / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    return input_dir


def user_templates_dir() -> Path:
    """
    Directory for custom template files.

    Returns:
        Path to ~/Documents/manuscript2slides/templates/
    """
    templates_dir = user_base_dir() / "templates"
    templates_dir.mkdir(parents=True, exist_ok=True)
    return templates_dir

def user_configs_dir() -> Path:
    """
    Directory for saved configuration files.

    Returns:
        Path to ~/Documents/manuscript2slides/configs/
    """
    configs_dir = user_base_dir() / "configs"
    configs_dir.mkdir(parents=True, exist_ok=True)
    return configs_dir

def user_manifests_dir() -> Path:
    """
    Directory for saved manifest files.

    Returns:
        Path to ~/Documents/manuscript2slides/manifests/
    """
    configs_dir = user_base_dir() / "manifests"
    configs_dir.mkdir(parents=True, exist_ok=True)
    return configs_dir