"""TODO docstring"""
import re
import io
import platform
from src.docx2pptx import config
import sys
import inspect

# region Utils - Basic
def debug_print(msg: str | list[str]) -> None:
    """Basic debug printing function"""
    if config.DEBUG_MODE:
        caller = inspect.stack()[1].function
        print(f"DEBUG [{caller}]: {msg}")

def setup_console_encoding() -> None:
    """Configure UTF-8 encoding for Windows console to prevent UnicodeEncodeError when printing non-ASCII characters (like emojis)."""
    if platform.system() == "Windows":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


# endregion

# region sanitize xml
def sanitize_xml_text(text: str) -> str:
    """Remove characters that aren't valid in XML."""
    if not text:
        return ""

    # Remove NULL bytes and control characters (except tab, newline, carriage return)
    sanitized = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)

    # Ensure it's a proper string
    return str(sanitized)
# endregion