# Building Binaries

This guide is for folks who want to build new standalone executables (OS-specific binaries) for manuscript2slides.

## Windows Binary

Creates a folder containing `manuscript2slides.exe` and all dependencies (Python runtime, Qt libraries, etc.). Nuitka compiles Python to C, then compiles the C into native machine code.

**Note**: We use `--standalone` mode (folder distribution) instead of `--onefile` (single .exe) because it triggers significantly fewer false positives from Windows Defender and other antivirus software. See [Nuitka issue #2495](https://github.com/Nuitka/Nuitka/issues/2495) for details.

### Requirements

- **Python**: 3.12 (recommended and tested)
  - Python 3.10-3.11 may work but are not tested
  - Python 3.13+ not yet supported by Nuitka 2.7.11
- **OS**: Windows 10 or later
- **Tools**:
  - Nuitka 2.7.11 (installed via pip) 
  - MSVC compiler (auto-installed by Nuitka on first build)

### Build Steps

```bash
# 1. Ensure you're using Python 3.12
python --version  # Should show 3.12.x

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

- **Location**: `deploy/gui.dist/` folder
- **Main executable**: `deploy/gui.dist/manuscript2slides.exe`
- **Size**: ~80-120 MB total (includes Python runtime, Qt libraries, and all dependencies)
- **Portable**: Yes - the entire folder can be copied to other Windows machines without Python installed

### Distribution

To prepare for distribution:

1. Rename the folder for clarity:
   ```powershell
   # In deploy/ directory
   Rename-Item gui.dist manuscript2slides
   ```

2. ZIP the folder:
   ```powershell
   Compress-Archive -Path manuscript2slides -DestinationPath manuscript2slides-windows.zip
   ```

The GitHub Actions workflow does this automatically on release tags.

### Build Script

The repository includes [build.py](../build.py), which wraps the Nuitka command:

```python
python build.py
```

This is the recommended way to build. For reference, here's what it does:

```bash
python -m nuitka \
  --standalone \
  --enable-plugin=pyside6 \
  --include-package-data=pptx \
  --include-package-data=docx \
  --include-package-data=manuscript2slides \
  --noinclude-qt-translations \
  --assume-yes-for-downloads \
  --windows-console-mode=disable \
  --output-dir=deploy \
  src\manuscript2slides\gui.py
```

### Testing the Build

After building, test the executable:

1. **Launch test**: Double-click `deploy\gui.dist\manuscript2slides.exe` - GUI should open without console window
2. **Smoke test**: Run a full conversion workflow:
   - File > Open > Select a .docx file
   - Click "Convert to Slides"
   - Verify output .pptx is created successfully
3. **Portability test**: Copy the entire `gui.dist` folder to a different machine (or clean VM) without Python and verify it works

### Troubleshooting

#### Build Errors

**"Python 3.13/3.14 not supported"**
- You're using Python 3.13 or 3.14, which aren't supported by Nuitka 2.7.11
- Switch to Python 3.12: `py -3.12 -m venv .venv`

**"ModuleNotFoundError: PySide6"**
- Dependencies not installed in current venv
- Run: `pip install -e .`

**"Missing file" errors at runtime**
- A library needs data files that weren't bundled
- Add `--include-package-data=<package_name>` flag to [build.py](../build.py) and rebuild

**Windows Defender blocks build**
- Antivirus may block Nuitka's resource bundling
- Add build folders to Windows Defender exclusions temporarily

<details>
<summary>How to add Windows Defender exclusions</summary>

1. **Open Windows Security**
   - Start > "Windows Security" > "Virus & threat protection"

2. **Add exclusions**
   - Scroll to "Virus & threat protection settings" > "Manage settings"
   - Scroll down to "Exclusions" > "Add or remove exclusions"
   - Click "Add an exclusion" > "Folder"

3. **Add these folders:**
   ```
   C:\Users\<YourUsername>\dev\manuscript2slides
   C:\Users\<YourUsername>\AppData\Local\Nuitka
   ```

4. **Try build again** - Windows Defender should no longer block the build process

</details>

#### Runtime Errors

**"No module named 'io'" or similar stdlib errors**
- Don't name your modules after Python stdlib modules (io, sys, os, etc.)

**Missing templates (pptx/docx)**
- Ensure `--include-package-data=pptx` and `--include-package-data=docx` flags are present in [build.py](../build.py)
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
- **Console mode**: Use `--windows-console-mode=disable` to prevent a console window from appearing on launch
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
