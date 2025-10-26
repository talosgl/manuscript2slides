## Package Dependencies
```bash
pip install -e '.[dev]'
```

## System Dependencies
On Windows, you'll need to install Python from python.org. That'll include Python and the UI library Tkinter.

Macs have Python preinstalled, and the system Python should have Tkinter already, but if you have trouble you can install python from python.org.

Linux usually has Python preinstalled, too, but Tkinter is excluded from some distros' versions. You'll need to install this package globally, rather than pip install it.

### Ubuntu/Debian
```bash
sudo apt install python3-tk
```

### Fedora/RHEL
```bash
sudo dnf install python3-tkinter
```
