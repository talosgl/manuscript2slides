"""Build script for creating binary executables using Nuitka. Uses platform detection to determine whether to output Mac or Windows binary."""

import subprocess
import sys
from pathlib import Path


def platform_helper() -> tuple:
    """Return platform-specific flags for binary builds for macos or windows."""
    if sys.platform == "darwin":
        return ("--macos-create-app-bundle", "--macos-app-icon=none") # TODO: Remove second flag if/when we add a cross-platform icon for the app.
    else:
        return ("--windows-console-mode=disable",)
    

def build() -> int:
    """Run Nuitka build command."""
    print("Building manuscript2slides...")
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
        *platform_helper(),
        "--output-dir=deploy",
        "--output-filename=manuscript2slides",
        str(Path("src") / "manuscript2slides" / "gui.py"),
    ]

    # Run build
    result = subprocess.run(cmd, encoding="utf-8", errors="replace")

    if result.returncode == 0:
        print("\nPASS Build successful!")
        ext = "exe"
        if sys.platform == "darwin":
            ext = "app"

        print(f"Output: {Path('deploy') / 'gui.dist' / f'manuscript2slides.{ext}'}")
        print("\nTo distribute: ZIP the entire 'deploy/gui.dist' folder")
    else:
        print("\nFAIL Build failed!")
        print("Check the output above for errors.")

    return result.returncode


if __name__ == "__main__":
    sys.exit(build())
