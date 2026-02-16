# Release Workflow

This guide covers:
- publishing new new manuscript2slides wheels to PyPI (`pip install` package site/host/distributor) 
- publishing new sets of binaries (Windows/Mac) via GitHub Releases

## Overview
manuscript2slides is available on macOS, Windows, and Linux. All platforms can use `pip install manuscript2slides` if the user already has Python installed and is comfortable with pip packages. To update the version available for pip, follow the [PyPI Distribution](#pypi-distribution) section.

Separately, we publish packaged binaries for Windows and macOS so that users do not need to know anything about Python. These are hosted on & downloadable via the Releases page for the repository on GitHub. To update these, follow the [Binary Builds](#packaged-binary-builds-for-windows--mac) section.

## PyPI Distribution

Publishing the Python package to PyPI allows users to install with `pip install manuscript2slides`.


### Re-publish to PyPI (Production)

#TODO: add short republish steps here.

---

## Packaged Binary Builds for Windows & Mac 

For major changes, like Python version upgrades, it is best to build the binaries locally first (on Windows and Mac machines) and smoke test them. To do that, follow the [making-binaries.md](/docs/making-binaries.md) guide. This guide will assume the current state of the repository's already been tested, and you're ready to trigger a new automated build to be released on GitHub.

The repository includes a GitHub Actions [workflow](../.github/workflows/binary-release.yml) that automatically builds Windows and macOS binaries when you push a version tag.

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
gh workflow run binary-release.yml -f version=0.2.0-test
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

