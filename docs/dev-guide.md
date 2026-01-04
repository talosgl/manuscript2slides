
## Development

### Install from Source

```bash
git clone https://github.com/talosgl/manuscript2slides.git # or your fork URL
cd manuscript2slides
pip install -e ".[dev]"
```

### Run Tests

```bash
pytest
```

### Run from Source

```bash
# Launch GUI from terminal
python -m manuscript2slides
python -m manuscript2slides.gui

# Launch CLI from terminal (from source directory)
python -m manuscript2slides.cli --input-docx="input\sample_doc.docx"
python -m manuscript2slides.cli --input-pptx="src\manuscript2slides\resources\sample_slides_output.pptx"
```

If you're using VS Code, you can use the `launch.json` provided by the repo to run more easily.

### Useful Commands
```bash
# From a venv terminal with [dev] installed

# General linting & fix
ruff check --fix

# Format code per repo standards, including import organization
ruff check --select I --fix && ruff format

# Check for type issues
mypy .

```

## Contributing

Contributions are welcome. Please feel free to:

- Report bugs via [GitHub Issues](https://github.com/talosgl/manuscript2slides/issues)
- Suggest features
- Submit pull requests

For substantial changes, please open an issue first to discuss your approach.

## Package Dependencies
```bash
pip install -e '.[dev]'
```

## System Dependencies
On Windows, you'll need to install Python from python.org. That'll include Python and pip.

Macs have Python preinstalled, but we recommend installing from python.org for the latest version.

Linux usually has Python preinstalled, too. However, you may need to install some system packages for Qt dependencies.

### Ubuntu/Debian
```bash
sudo apt install libxcb-xinerama0 libxcb-cursor0
```

### Fedora/RHEL
```bash
sudo dnf install xcb-util-cursor
```

## Troubleshooting

### Qt platform plugin errors on Linux
If you get errors about "qt.qpa.plugin", install the system packages above.

### "No module named 'PySide6'" error
Run `pip install -e '.[dev]'` to install all dependencies.