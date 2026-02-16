# Build Package for PyPI

This doc covers building wheels for PyPI and the *first-time* release process.

### Requirements

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
