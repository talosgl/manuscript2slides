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

The repository includes [build.py](../build.py), which wraps the Nuitka command with platform detection:

```python
python build.py
```

This is the recommended way to build. The script automatically detects the platform and applies the correct flags (e.g., `--macos-create-app-bundle` on macOS, `--windows-console-mode=disable` on Windows).

It also references [nuitka-package.config.yaml](../nuitka-package.config.yaml), which patches python-pptx and python-docx template path resolution for compatibility with compiled builds on macOS/Linux. See [Template path resolution](#template-path-resolution-macos) for details.

### Testing the Build

After building, run through the full smoke test checklist in [manual-smoke-test.md](manual-smoke-test.md) to verify the binary works correctly.

### Troubleshooting

#### Build Errors

**"Python 3.13/3.14 not supported"**
- You're using Python 3.13 or 3.14, which aren't supported by Nuitka 2.7.11
- Switch to Python 3.12: `py -3.12 -m venv .venv` (Windows) or `python3.12 -m venv .venv` (macOS)

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

#### Template path resolution (macOS)

python-pptx and python-docx resolve template files using relative `..` paths (e.g., `pptx/oxml/../templates/notes.xml`). Nuitka compiles away the Python source directories, so on macOS/Linux the `..` traversal fails because the intermediate directory doesn't exist. On Windows, the OS normalizes paths before access, so this isn't an issue.

The fix is in [nuitka-package.config.yaml](../nuitka-package.config.yaml), which uses Nuitka's `replacements_plain` feature to wrap these path constructions with `os.path.normpath()`, resolving the path before file access.

If you see `[Errno 2] No such file or directory` errors referencing `../templates/` in the macOS build, check whether a new version of python-pptx or python-docx has added more `..` template lookups, and add corresponding entries to the YAML config.

### Build Artifacts

The following directories are created during builds and should be in `.gitignore`:

```
deploy/                 # Final binary output (.exe or .app)
*.build/                # Nuitka build cache
*.dist/                 # Nuitka distribution files (Windows)
*.onefile-build/        # Nuitka onefile temp directory
```

### Notes

- **Why not pyside6-deploy?** We tried it first, but it ignored config file settings. Direct Nuitka gives us full control.
- **Onefile vs Standalone mode**: `--onefile` creates a single file (easier distribution). `--standalone` creates a folder/bundle (faster startup). We use `--standalone` on both platforms.
- **Console mode** (Windows): `--windows-console-mode=disable` prevents a console window from appearing on launch
- **Code signing**: Not currently implemented for either platform. Windows users will see a SmartScreen warning; macOS users need to right-click > Open on first launch. See the [macOS Binary](#macos-binary) section for details. If this is a barrier for you, [open an issue](https://github.com/talosgl/manuscript2slides/issues).

---

## macOS Binary

Creates a `.app` bundle containing the compiled application and all dependencies. Nuitka compiles Python to C, then compiles the C into native machine code, and wraps it in a standard macOS application bundle.

### Requirements

- **Python**: 3.12 (recommended and tested)
  - Python 3.10-3.11 may work but are not tested
  - Python 3.13+ not yet supported by Nuitka 2.7.11
- **OS**: macOS 15.0 or later (Apple Silicon only; Intel Mac support can be added if requested - [open an issue](https://github.com/talosgl/manuscript2slides/issues))
- **Tools**:
  - Nuitka 2.7.11 (installed via pip)
  - Xcode Command Line Tools (`xcode-select --install`)

### Build Steps

```bash
# 1. Ensure you're using Python 3.12
python3.12 --version  # Should show 3.12.x

# 2. Create/activate venv (if needed)
python3.12 -m venv .venv
source .venv/bin/activate

# 3. Install dependencies
pip install -e .
pip install Nuitka==2.7.11

# 4. Run the build
python build.py
```

**Build time**: ~15-20 minutes on first build, faster on subsequent builds with ccache.

### Output

- **Location**: `deploy/gui.app`
- **Size**: ~100 MB (includes Python runtime, Qt libraries, and all dependencies)
- **Portable**: Yes - the `.app` can be copied to other Apple Silicon Macs without Python installed

### Distribution

To prepare for distribution:

1. Rename the app bundle:
   ```bash
   mv deploy/gui.app deploy/manuscript2slides.app
   ```

2. ZIP it:
   ```bash
   cd deploy
   zip -r manuscript2slides-macos.zip manuscript2slides.app
   ```

The GitHub Actions workflow does this automatically on release tags.

### First Launch (unsigned app)

Since the app is not signed or notarized:

1. Right-click `manuscript2slides.app` and select "Open" (required for unsigned apps)
2. macOS will warn the app is from an unidentified developer. To proceed:
   1. Click the upper-right question mark for the Apple help page on the topic.
   2. Click "Done" in the prompt to dismiss it.
   3. Open your macOS System Settings, go to Privacy & Security, and scroll down to Security.
   4. Note the message: `"manuscript2slides" was blocked to protect your Mac.`
   5. Click "Open Anyway"; input your user account info if prompted.
   6. The app should now run normally, and should not prompt you again on subsequent launches unless you get a new version.
3. When prompted to access your Documents folder, click "Allow" - the app stores its templates and sample files there
4. After the first launch, you can open it normally by double-clicking

### Troubleshooting

See the shared [Troubleshooting](#troubleshooting) section above. macOS-specific issues:

**"No such file or directory" referencing `../templates/`**
- This is the template path resolution issue described [above](#template-path-resolution-macos)
- Check [nuitka-package.config.yaml](../nuitka-package.config.yaml) for missing entries

**App doesn't open at all (no error)**
- Try launching from terminal to see error output: `./deploy/gui.app/Contents/MacOS/gui`

**Gatekeeper blocks the app entirely**
- Use right-click > Open instead of double-clicking
- Or run: `xattr -cr deploy/manuscript2slides.app` to remove quarantine attributes
