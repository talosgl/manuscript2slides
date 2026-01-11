# Build Guide

This document describes how to build and distribute manuscript2slides.

## Table of Contents

- [PyPI Distribution](#pypi-distribution)
- [Binary Builds](#binary-builds)
  - [Windows Binary](#windows-binary)
  - [macOS Binary](#macos-binary-coming-soon)

---

## PyPI (pip) Distribution

Building and publishing the Python package to PyPI allows users to install with `pip install manuscript2slides`.

### Prerequisites

- Python 3.10+
- [build](https://pypi.org/project/build/) and [twine](https://pypi.org/project/twine/) installed

### Build the Package

```bash
# Install build tools
pip install build twine

# Build source distribution and wheel
python -m build

# Output will be in dist/:
# - manuscript2slides-0.1.5.tar.gz (source distribution)
# - manuscript2slides-0.1.5-py3-none-any.whl (wheel)
```

### Test the Build Locally

```bash
# Create a fresh venv for testing
python -m venv test-venv
test-venv\Scripts\activate     # Windows
source test-venv/bin/activate  # macOS/Linux

# Install from the wheel
pip install dist/manuscript2slides-0.1.5-py3-none-any.whl

# Test it works
python -m manuscript2slides # Ought to launch GUI
python -m manuscript2slides-cli --help
```

### Publish to TestPyPI (for testing)

```bash
# Upload to TestPyPI
python -m twine upload --repository testpypi dist/*

# Install from TestPyPI to verify
pip install --index-url https://test.pypi.org/simple/ manuscript2slides
```

### Publish to PyPI (production)

```bash
# Upload to real PyPI
python -m twine upload dist/*

# Users can now install with:
# pip install manuscript2slides
```

**Note**: You'll need PyPI credentials. Set up an API token at https://pypi.org/manage/account/token/

---

## Binary Builds

Standalone executable builds that include Python runtime and all dependencies. Users don't need Python installed.

---

## Windows Binary

Creates a standalone `.exe` file for Windows.

### Requirements

- **Python**: 3.10, 3.11, or 3.12 (Python 3.13+ not yet supported by Nuitka 2.7.11)
- **OS**: Windows 10 or later
- **Tools**:
  - Nuitka 2.7.11 (installed via pip)
  - MSVC compiler (auto-installed by Nuitka on first build)

### Quick Start

```bash
# 1. Ensure you're using Python 3.10-3.12
python --version  # Should show 3.10.x, 3.11.x, or 3.12.x

# 2. Create/activate venv (if needed)
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -e .
pip install Nuitka==2.7.11

# 4. Run the build
python build.py
```

**Build time**: ~15-20 minutes on first build, ~10-15 minutes on subsequent builds.

### Output

- **Location**: `deploy/manuscript2slides.exe`
- **Size**: ~80-120 MB (includes Python runtime, Qt libraries, and all dependencies)
- **Portable**: Yes - can be copied to other Windows machines without Python installed

### Build Script

The repository includes [build.py](build.py), which wraps the Nuitka command for convenience:

```python
python build.py
```

This is the recommended way to build. It's easier to type and works consistently across local builds and CI/CD (GitHub Actions).

### Build Command Explained

For reference, here's what `build.py` does under the hood:

```bash
python -m nuitka \
  --onefile \                                    # Single .exe file (not a folder)
  --enable-plugin=pyside6 \                      # Qt plugin support
  --include-package-data=pptx \                  # Bundle python-pptx templates
  --include-package-data=docx \                  # Bundle python-docx templates
  --include-package-data=manuscript2slides \     # Bundle app resources
  --noinclude-qt-translations \                  # Skip Qt translation files (reduces size)
  --assume-yes-for-downloads \                   # Auto-download build tools
  --windows-console-mode=disable \               # No console window (GUI app)
  --output-dir=deploy \                          # Output directory
  --output-filename=manuscript2slides.exe \      # Output filename
  src\manuscript2slides\gui.py                   # Entry point
```

### Testing the Build

After building, test the executable:

1. **Launch test**: Double-click `deploy\manuscript2slides.exe` - GUI should open without console window
2. **Smoke test**: Run a full conversion workflow:
   - File > Open > Select a .docx file
   - Click "Convert to Slides"
   - Verify output .pptx is created successfully
3. **Portability test**: Copy `.exe` to a different machine (or clean VM) without Python and verify it works

### Automated Builds (GitHub Actions)

The repository includes a GitHub Actions workflow that automatically builds Windows binaries when you push a version tag.

#### Creating a Release

```bash
# 1. Commit all your changes
git add .
git commit -m "Your commit message"

# 2. Create an annotated version tag
git tag -a v0.1.5 -m "Release v0.1.5: Description of changes"

# 3. Push the tag to GitHub
git push origin v0.1.5
```

#### What Happens Next

1. GitHub Actions detects the `v*.*.*` tag
2. Spins up a Windows runner with Python 3.12
3. Installs dependencies and Nuitka 2.7.11
4. Runs `python build.py`
5. Creates a GitHub Release with `manuscript2slides.exe` attached
6. Release includes installation instructions

**Monitoring**: Watch build progress at `https://github.com/talosgl/manuscript2slides/actions`

**Build time**: ~15-20 minutes

**If the build fails**:
- Check the Actions tab for error logs
- Common issues: missing dependencies, Python version mismatch
- Test the build locally first with `python build.py` before pushing tags

**Workflow file**: [.github/workflows/build-release.yml](.github/workflows/build-release.yml)

### Troubleshooting

#### Build Errors

**"Python 3.13/3.14 not supported"**
- You're using Python 3.13 or 3.14, which aren't supported by Nuitka 2.7.11
- Switch to Python 3.10-3.12: `py -3.12 -m venv .venv`

**"ModuleNotFoundError: PySide6"**
- Dependencies not installed in current venv
- Run: `pip install -e .`

**"Missing file" errors at runtime**
- A library needs data files that weren't bundled
- Add `--include-package-data=<package_name>` flag and rebuild

**Windows Defender blocks build**
- Antivirus may block Nuitka's resource bundling
- Add build folders to Windows Defender exclusions temporarily

#### Runtime Errors

**Missing templates (pptx/docx)**
- Ensure `--include-package-data=pptx` and `--include-package-data=docx` flags are present
- These bundle the XML templates that python-pptx and python-docx need

### Build Artifacts

The following directories are created during builds and should be in `.gitignore`:

```
deploy/                 # Final .exe output
*.build/                # Nuitka build cache
*.dist/                 # Nuitka distribution files
*.onefile-build/        # Nuitka onefile temp directory
```

### Notes

- **Why not pyside6-deploy?** We tried it first, but it ignored config file settings. Direct Nuitka gives us full control.
- **Onefile vs Standalone mode**: `--onefile` creates a single .exe (easier distribution). `--standalone` creates a folder with .exe + DLLs (faster startup, harder to distribute).
- **Console mode**: GUI apps should use `--windows-console-mode=disable` to prevent a console window from appearing.
- **Code signing**: Not planned - costs money and requires annual renewal. Users will see "Unknown publisher" warning, but the app will still run fine if they click "More info" > "Run anyway".

---

## macOS Binary (Coming Soon)

Building standalone `.app` bundles for macOS will follow a similar approach using Nuitka.

**Planned approach**:
- Use Nuitka with macOS-specific flags
- Create `.app` bundle structure
- Test on multiple macOS versions
- Set up GitHub Actions workflow for automated builds

**Documentation will be added here once implemented.**
