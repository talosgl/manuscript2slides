# Building manuscript2slides

Quick reference for locally building pip packages and platform-specific binaries. See [docs/](docs/) for detailed documentation.


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

See [docs/making-binaries.md](docs/making-binaries.md) for detailed instructions and troubleshooting.

## Automated Binary Releases

Push a version tag to trigger automated builds:

```bash
git tag -a v0.2.0 -m "Release v0.2.0: Description"
git push origin v0.2.0
```

GitHub Actions will build both Windows and macOS binaries and create a release automatically.

### Draft Binary Build Release

To test the build process without creating a public release, use manual workflow dispatch:

```bash
gh workflow run binary-release.yml -f version=0.2.0-test
```