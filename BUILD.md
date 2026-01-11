# Building Windows Binaries

This document describes how to build standalone `.exe` binaries for manuscript2slides on Windows.

## Requirements

- **Python**: 3.10, 3.11, or 3.12 (Python 3.13+ not yet supported by Nuitka 2.7.11) from https://www.python.org/downloads/windows/
- **OS**: Windows 11
- **Tools**:
  - Nuitka 2.7.11 (installed via pip)
  - MSVC compiler (auto-installed by Nuitka on first build)

## Quick Start

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

# Or run Nuitka directly:
# python -m nuitka --onefile --enable-plugin=pyside6 --include-package-data=pptx --include-package-data=docx --include-package-data=manuscript2slides --noinclude-qt-translations --assume-yes-for-downloads --windows-console-mode=disable --output-dir=deploy --output-filename=manuscript2slides.exe 'src\manuscript2slides\gui.py'
```

Build time: ~15-20 minutes on first build, ~10-15 minutes on subsequent builds.

## Output

- **Location**: `deploy/manuscript2slides.exe`
- **Size**: ~80-120 MB (includes Python runtime, Qt libraries, and all dependencies)
- **Portable**: Yes - can be copied to other Windows machines without Python installed

## Build Script

The repository includes `build.py`, which wraps the Nuitka command for convenience:

```python
python build.py
```

This is the recommended way to build. It's easier to type and works consistently across local builds and CI/CD (GitHub Actions).

## Build Command Explained (what build.py does)

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

## Testing the Build

After building, test the executable:

1. **Launch test**: Double-click `deploy\manuscript2slides.exe` - GUI should open without console window
2. **Smoke test**: Run a full conversion workflow:
   - File > Open > Select a .docx file
   - Click "Convert to Slides"
   - Verify output .pptx is created successfully
3. **Portability test**: Copy `.exe` to a different machine (or clean VM) without Python and verify it works

## Troubleshooting

### Build Errors

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

### Runtime Errors

**"No module named 'io'" or similar stdlib errors**
- This was caused by having a file named `io.py` in the source code
- Don't name your modules after Python stdlib modules

**Missing templates (pptx/docx)**
- Ensure `--include-package-data=pptx` and `--include-package-data=docx` flags are present
- These bundle the XML templates that python-pptx and python-docx need

## Build Artifacts (Gitignored)

The following directories are created during builds and should be in `.gitignore`:

```
deploy/                 # Final .exe output
*.build/                # Nuitka build cache
*.dist/                 # Nuitka distribution files
*.onefile-build/        # Nuitka onefile temp directory
```

## Notes

- **Why not pyside6-deploy?** We tried it first, but it ignored config file settings. Direct Nuitka gives us full control.
- **Onefile vs Standalone mode**: `--onefile` creates a single .exe (easier distribution). `--standalone` creates a folder with .exe + DLLs (faster startup, harder to distribute).
- **Console mode**: GUI apps should use `--windows-console-mode=disable` to prevent a console window from appearing.

## Future Improvements

- Add custom application icon (currently uses default)
- macOS .app bundle builds

## Not Planning
- Code signing for Windows SmartScreen - Costs money and requires annual renewal. Users will see "Unknown publisher" warning, but the app will still run fine if they click "More info" > "Run anyway".
