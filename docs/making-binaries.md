# Making Binaries

Detailed guide for making new standalone executables (OS-specific binaries for macOS and Windows) for manuscript2slides.

## Requirements & Info
- **Python**: 3.12 (recommended and tested)
- Nuitka 2.7.11 (see [requirements-binary-build.txt](/requirements-binary-build.txt))
- **OS**: Windows 10 or later; macOS 15.0 or later (Apple Silicon only; Intel Mac support can be added if requested - [open an issue](https://github.com/talosgl/manuscript2slides/issues))
- **Tools**:
  - (Windows) MSVC compiler (should be auto-installed by Nuitka on first build)
  - (Mac) Xcode Command Line Tools (`xcode-select --install`)

## Build Steps

```bash
# 1. Ensure you're using Python 3.12
python --version  # Should show 3.12.x

# 2. Create/activate venv (if needed)
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -e .
pip install Nuitka==2.7.11 # or let your venv setup tool use requirements-binary-build.txt to get it

# 4. Run the build
python make_binary.py
```

## Windows Binary Info

Creates a folder containing `manuscript2slides.exe` and all dependencies (Python runtime, Qt libraries, etc.). Nuitka compiles Python to C, then compiles the C into native machine code.

Part of why we use `--standalone` mode (folder distribution) instead of `--onefile` (single .exe) is because it triggers significantly fewer false positives from Windows Defender and other antivirus software. See [Nuitka issue #2495](https://github.com/Nuitka/Nuitka/issues/2495) for details.

### Output

- **Location**: `deploy/gui.dist/` folder
- **Main executable**: `deploy/gui.dist/manuscript2slides.exe`
- **Size**: ~80-120 MB total (includes Python runtime, Qt libraries, and all dependencies)
- **Portable**: Yes - the entire folder can be copied to other Windows machines without Python installed


## macOS Binary Info

Creates a `.app` bundle containing the compiled application and all dependencies. Nuitka compiles Python to C, then compiles the C into native machine code, and wraps it in a standard macOS application bundle.

### Output

- **Location**: `deploy/gui.app`
- **Size**: ~100 MB (includes Python runtime, Qt libraries, and all dependencies)
- **Portable**: Yes - the `.app` can be copied to other Apple Silicon Macs without Python installed

### Known Issue: Template path resolution (macOS)

python-pptx and python-docx resolve template files using relative `..` paths (e.g., `pptx/oxml/../templates/notes.xml`). Nuitka compiles away the Python source directories, so on macOS/Linux the `..` traversal fails because the intermediate directory doesn't exist. On Windows, the OS normalizes paths before access, so this isn't an issue.

The fix is in [nuitka-package.config.yaml](../nuitka-package.config.yaml), which uses Nuitka's `replacements_plain` feature to wrap these path constructions with `os.path.normpath()`, resolving the path before file access.

If you see `[Errno 2] No such file or directory` errors referencing `../templates/` in the macOS build, check whether a new version of python-pptx or python-docx has added more `..` template lookups, and add corresponding entries to the YAML config.

## Troubleshooting Build Errors

**"Python 3.13/3.14 not supported"**
- You're using Python 3.13 or 3.14, which aren't supported by Nuitka 2.7.11
- Switch to Python 3.12: `py -3.12 -m venv .venv` (Windows) or `python3.12 -m venv .venv` (macOS)

**"ModuleNotFoundError: PySide6"**
- Dependencies not installed in current venv
- Run: `pip install -e .`

**"Missing file" errors at runtime**
- A library needs data files that weren't bundled
- Add `--include-package-data=<package_name>` flag to [make_binary.py](../make_binary.py) and rebuild
- If on Mac, check [the known issues section](#known-issue-template-path-resolution-macos)

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


### Notes

- **Why not pyside6-deploy?** We tried it first, but it ignored config file settings. Direct Nuitka gives us full control.
- **Onefile vs Standalone mode**: We use `--standalone` on both platforms.
- **Console mode** (Windows): `--windows-console-mode=disable` prevents a console window from appearing on launch. Unnecessary on Mac.
- **Code signing**: Not currently implemented for either platform. Windows users will see a SmartScreen warning; macOS users need to right-click > Open on first launch. See the [macOS Binary](#macos-binary) section for details. If this is a barrier for you, [open an issue](https://github.com/talosgl/manuscript2slides/issues).

---

## About make_binary Script

[make_binary.py](../make_binary.py) wraps the Nuitka command with platform detection and applies the correct flags (e.g., `--macos-create-app-bundle` on macOS, `--windows-console-mode=disable` on Windows).

It also references [nuitka-package.config.yaml](../nuitka-package.config.yaml), which patches python-pptx and python-docx template path resolution for compatibility with compiled builds on macOS/Linux. See [Template path resolution](#known-issue-template-path-resolution-macos) for details.

## Testing the Build

After building, run through the full smoke test checklist in [manual-smoke-test.md](manual-smoke-test.md) to verify the binary works correctly.


## Distribution
We don't distribute manually/locally build binaries, we use GitHub Actions to output to the releases page of the repo. One thing we'd need to do, if we *were* going to distribute local output, is zip up the output of the script and name it appropriately. But the GitHub Actions workflow does this automatically for both platforms. See [releasing](/docs/releasing.md).


