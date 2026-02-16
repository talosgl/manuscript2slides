# Release Workflow

This guide covers:
- publishing new new manuscript2slides wheels to PyPI (`pip install` package site/host/distributor) 
- publishing new sets of binaries (Windows/Mac) via GitHub Releases

## Overview
manuscript2slides is available on macOS, Windows, and Linux. All platforms can use `pip install manuscript2slides` if the user already has Python installed and is comfortable with pip packages. To update the version available for pip, follow the [PyPI Distribution](#pypi-distribution) section.

Separately, we publish packaged binaries for Windows and macOS so that users do not need to know anything about Python. These are hosted on & downloadable via the Releases page for the repository on GitHub. To update these, follow the [Binary Builds](#packaged-binary-builds-for-windows--mac) section.

## PyPI Distribution

Publishing the Python package to PyPI allows users to install with `pip install manuscript2slides`.

### Prerequisites

- Python 3.10+
- [build](https://pypi.org/project/build/) and [twine](https://pypi.org/project/twine/) installed
- TestPyPI and PyPI accounts with API tokens

### First-Time Setup: PyPI Accounts

You need two separate accounts:

1. **TestPyPI** (for testing): https://test.pypi.org/account/register/
2. **Real PyPI** (for production): https://pypi.org/account/register/

For both accounts:
- Register and verify email
- **Set up 2FA** (required for PyPI)
- Create an API token (Settings > API tokens > "Add API token")
  - **Note**: Start with an account-scoped token for initial upload
  - After first successful upload, rescope to project-specific token (see below)

Save these tokens securely - you'll use them instead of passwords.

### Configure Credentials

Option 1: Use `.pypirc` file. You can save just the username and not the password, if you prefer.

Create `~/.pypirc` (Linux/Mac) or `%USERPROFILE%\.pypirc` (Windows):

```ini
[pypi]
username = __token__
password = pypi-...  # Your PyPI token here

[testpypi]
username = __token__
password = pypi-...  # Your TestPyPI token
```

Option 2: Enter credentials when prompted by `twine upload`

### Build the Package

```bash
# Install build tools (if not already installed)
pip install build twine

# Build source distribution and wheel
python -m build

# (Optional) Check package contents
python -m zipfile -l dist/manuscript2slides-<version>-py3-none-any.whl
```

**Output** (in `dist/`):
- `manuscript2slides-<version>.tar.gz` (source distribution)
- `manuscript2slides-<version>-py3-none-any.whl` (wheel - what pip actually installs)

### Test the Build Locally

```bash
# Create a fresh venv for testing
python -m venv test-venv
test-venv\Scripts\activate     # Windows
source test-venv/bin/activate  # macOS/Linux

# Install from the wheel
pip install dist/manuscript2slides-<version>-py3-none-any.whl

# Test it works
python -m manuscript2slides  # Should launch GUI
```

### Publish to TestPyPI

Always test on TestPyPI first:

```bash
# Upload to TestPyPI
python -m twine upload --repository testpypi dist/*
```

If you didn't set up `.pypirc`, it will prompt for:
- Username: `__token__`
- Password: Your TestPyPI API token (starts with `pypi-...`)

**Wait 2-5 minutes** for the package to propagate on TestPyPI servers.

### Test Install from TestPyPI

```bash
# Create fresh venv
python -m venv test-venv
test-venv\Scripts\activate     # Windows
source test-venv/bin/activate  # macOS/Linux

# Install from TestPyPI
# --extra-index-url gets dependencies from real PyPI
pip install --no-cache-dir \
  --index-url https://test.pypi.org/simple/ \
  --extra-index-url https://pypi.org/simple/ \
  manuscript2slides

# Test it works
python -m manuscript2slides  # Should launch GUI
```

**Flags explained:**
- `--no-cache-dir`: Force fresh download (important for testing new versions)
- `--index-url`: Get your package from TestPyPI
- `--extra-index-url`: Get dependencies (PySide6, etc.) from real PyPI

### Fix and Reupload (if needed)

If you find issues during TestPyPI testing and need to re-upload:

```bash
# 1. Bump version in pyproject.toml (e.g., 0.1.4 -> 0.1.5)
#    Note: TestPyPI versions don't need to match real PyPI

# 2. Delete old build
rm -r dist/  # Linux/Mac/Git Bash
# OR
rmdir /s dist  # Windows CMD

# 3. Rebuild
python -m build

# 4. Reupload to TestPyPI
python -m twine upload --repository testpypi dist/*

# 5. Wait 2-5 minutes, then retest with --no-cache-dir
```

Repeat until it works perfectly on TestPyPI.

### Publish to PyPI (Production)

Once verified on TestPyPI:

```bash
# Upload to real PyPI
python -m twine upload dist/*

# Users can now install with:
# pip install manuscript2slides
```

**Wait 2-5 minutes** for propagation, then test:

```bash
pip install --no-cache-dir manuscript2slides
python -m manuscript2slides
```

### Post-Release: Rescope PyPI API Tokens (Recommended)

After your first successful upload to PyPI, improve security by rescoping tokens:

1. Go to TestPyPI & PyPI settings
2. Create new tokens scoped to `manuscript2slides` only (not account-wide)
3. Delete the account-wide tokens
4. Update `~/.pypirc` with the new project-scoped tokens

**Why:** Project-scoped tokens can only upload to `manuscript2slides`, limiting damage if leaked.

---

## Packaged Binary Builds for Windows & Mac 

For major changes, like Python version upgrades, it is best to build the binaries locally first (on Windows and Mac machines) and smoke test them. To do that, follow the [building.md](/building.md) guide. This guide will assume the current state of the repository's already been tested, and you're ready to trigger a new automated build to be releaseed on GitHub.

The repository includes a GitHub Actions [workflow](../.github/workflows/build-release.yml) that automatically builds Windows and macOS binaries when you push a version tag.

### Kick a new Release Build

```bash
# 1. Commit all your changes
git add .
git commit -m "Your commit message"

# 2. Create an annotated version tag
git tag -a v<version> -m "Release v<version>: Description of changes"

# 3. Push the tag to GitHub
git push origin v<version>
```

### What Happens Next

1. GitHub Actions detects the `v*.*.*` tag
2. **Runs tests first** (on Ubuntu with Python 3.12)
3. If tests pass, spins up Windows and macOS runners **in parallel**
4. Each runner installs dependencies, Nuitka 2.7.11, and runs `python make_binary.py`
5. Both runners upload their build artifacts
6. A final release job downloads both artifacts, computes checksums, and creates a GitHub Release with both ZIPs attached
7. Release includes platform-specific installation instructions

**If tests fail**, both builds are skipped and you'll get notified.

### Testing Builds (Draft Release)

You can test the full build pipeline without creating a public release using workflow dispatch. This creates a **draft** release (only visible to logged-in contributors).

**Option 1: GitHub Web UI**

1. Go to [Actions](https://github.com/talosgl/manuscript2slides/actions)
2. Click "Build Release" in the left sidebar
3. Click "Run workflow" dropdown (top right)
4. Enter a version string (e.g., `0.2.0-test`)
5. Click "Run workflow"

**Option 2: GitHub CLI**

```bash
gh workflow run build-release.yml -f version=0.2.0-test
```

**Accessing the draft release:**

1. Go to the repository's [Releases page](https://github.com/talosgl/manuscript2slides/releases)
2. Draft releases appear at the top with a "Draft" label
3. Click the release to download the attached ZIP files
4. Test the binaries locally (run through [manual-smoke-test.md](manual-smoke-test.md))

**Managing drafts:**
- **Publish**: Click "Edit" on the draft, then "Publish release" to make it visible to everyone
- **Delete**: Click "Edit", scroll down, then "Delete this release"

### Monitoring the Build

Watch build progress at: `https://github.com/talosgl/manuscript2slides/actions`

**Build time**: ~15-20 minutes per platform (both run in parallel)

**If the build fails**:
- Check the Actions tab for error logs
- Common issues: missing dependencies, Python version mismatch

---

## Version Numbering

We use [Semantic Versioning](https://semver.org/):

- **v0.1.0 -> v0.1.1**: Patch (bug fixes, small changes)
- **v0.1.0 -> v0.2.0**: Minor (new features, backwards compatible)
- **v0.1.0 -> v1.0.0**: Major (breaking changes, stable release)

