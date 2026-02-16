# Building manuscript2slides

Quick reference for building pip packages and platform-specific binaries. See [docs/](docs/) for detailed documentation.


## PyPI Package

```bash
# Build
python -m build

# Upload to TestPyPI
python -m twine upload --repository testpypi dist/*

# Upload to PyPI
python -m twine upload dist/*
```

See [docs/releasing.md](docs/releasing.md) for the full release process.

---

## Platform Binaries (Windows & macOS)

```bash
# Install dependencies. Assumes you already have Python 3.12
pip install -e .
pip install Nuitka==2.7.11

# Build (auto-detects platform)
python make_binary.py
```

**Local Output**:
- **Windows**: `deploy/gui.dist/` folder containing `manuscript2slides.exe` (~100MB total). 
- **macOS**: `deploy/gui.app` bundle (~100MB).

**Note**: We use standalone mode (folder/bundle distribution) instead of single-file to reduce Windows Defender false positives.

**Requirements**: Python 3.12, Windows 10+ or macOS 15.0+ (Apple Silicon)

See [docs/building.md](docs/building.md) for detailed instructions and troubleshooting.

## Automated Releases

Push a version tag to trigger automated builds:

```bash
git tag -a v0.2.0 -m "Release v0.2.0: Description"
git push origin v0.2.0
```

GitHub Actions will build both Windows and macOS binaries and create a release automatically.

### Test Binary Build Release

To test the build process without creating a public release, use manual workflow dispatch:

```bash
gh workflow run build-release.yml -f version=0.2.0-test
```

This creates a draft release (only visible to logged-in contributors). To access it:

1. Go to the repository's [Releases page](https://github.com/talosgl/manuscript2slides/releases)
2. Draft releases appear at the top with a "Draft" label
3. Click the release to download the attached ZIP files on each platform
4. Run [docs/manual-smoke-test.md](docs/manual-smoke-test.md) on each build, per platform, to verify

To publish or delete a draft:
- **Publish**: Click "Edit" on the draft, then "Publish release" to make it visible to everyone
- **Delete**: Click "Edit", scroll down, then "Delete this release"

See [docs/releasing.md](docs/releasing.md) for details.

