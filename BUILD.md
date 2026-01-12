# Building manuscript2slides

Quick reference for building binaries and packages. See [docs/](docs/) for detailed documentation.

## Windows Binary

```bash
# Install dependencies. Assumes you already have Python 3.12
pip install -e .
pip install Nuitka==2.7.11

# Build
python build.py
```

**Output**: `deploy/gui.dist/` folder containing `manuscript2slides.exe` and dependencies (~100MB total)

**For distribution**: Rename `gui.dist` to `manuscript2slides` before zipping for users

**Note**: We use standalone mode (folder distribution) instead of single-file to reduce Windows Defender false positives.

**Requirements**: Python 3.10-3.12, Windows 10+

See [docs/building.md](docs/building.md) for detailed instructions and troubleshooting.

## macOS Binary

Coming soon. See [docs/building.md](docs/building.md)

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

## Automated Releases

Push a version tag to trigger automated builds:

```bash
git tag -a v0.2.0 -m "Release v0.2.0: Description"
git push origin v0.2.0
```

GitHub Actions will build the Windows binary and create a release automatically.

See [docs/releasing.md](docs/releasing.md) for details.
