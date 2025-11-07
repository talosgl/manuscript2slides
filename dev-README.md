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