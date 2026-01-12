"""Build script for creating Windows binary executable using Nuitka."""

import subprocess
import sys
from pathlib import Path


def build() -> int:
    """Run Nuitka build command."""
    print("Building manuscript2slides.exe...")
    print("This will take 15-20 minutes on first build.\n")

    # Nuitka build command
    cmd = [
        sys.executable,  # Use the current Python interpreter
        "-m",
        "nuitka",
        "--standalone",  # Changed from --onefile to reduce antivirus false positives
        "--enable-plugin=pyside6",
        "--include-package-data=pptx",
        "--include-package-data=docx",
        "--include-package-data=manuscript2slides",
        "--noinclude-qt-translations",
        "--assume-yes-for-downloads",
        "--windows-console-mode=disable",
        "--output-dir=deploy",
        "--output-filename=manuscript2slides",
        str(Path("src") / "manuscript2slides" / "gui.py"),
    ]

    # Run build
    result = subprocess.run(cmd, encoding="utf-8", errors="replace")

    if result.returncode == 0:
        print("\nPASS Build successful!")
        print(f"Output: {Path('deploy') / 'gui.dist' / 'manuscript2slides.exe'}")
        print("\nTo distribute: ZIP the entire 'deploy/gui.dist' folder")
    else:
        print("\nFAIL Build failed!")
        print("Check the output above for errors.")

    return result.returncode


if __name__ == "__main__":
    sys.exit(build())
