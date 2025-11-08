
# manuscript2slides

Convert Microsoft Word documents to PowerPoint slides and back again.

## Features
Turn your Microsoft Word manuscripts into presentation slides (and back again), with a simple desktop interface.

- **Multiple chunking strategies**: By paragraph, page, or heading (flat/nested)
- **Formatting preservation**: Bold, italics, colors, highlights, strikethrough, super/subscript, and more
- **Annotation support**: Comments, footnotes, and endnotes can be copied to slide speaker notes
- **Round-trip capability**: Convert DOCX → PPTX → DOCX with optional metadata preservation
- **Cross-platform**: Works on Windows, macOS, and Linux
- **Both GUI and CLI**: Use whichever fits your workflow
- **pip-installable**: Usable as a Python library for scripted or automated conversions

## Quick Start Guide

### Option 1: Desktop App (recommended)

1. Download the latest release for your platform from the [latest release](https://github.com/talosgl/manuscript2slides/releases/latest) page.Just download, extract (if needed), and run. No installation required.
2. Open `manuscript2slides` (no install needed).
3. Choose a tab:
   - **DOCX → PPTX:** Convert your manuscript into slides.
   - **PPTX → DOCX:** Turn a slide deck back into text.
   - **DEMO:** Try a sample conversion.

4. Pick your file(s), adjust options (chunking, formatting, annotations), and click **Convert!**

### Option 2: Command Line (advanced)

You can install manuscript2slides via the executables above, or use pip:

```bash
pip install manuscript2slides
```

Then you can run it from the command line like:
```bash
# Opens GUI
manuscript2slides

# Convert a Word document
manuscript2slides --input-docx my-manuscript.docx

# Reverse conversion
manuscript2slides --input-pptx presentation.pptx --direction pptx2docx

# See a demo dry run with sample files
manuscript2slides --demo-round-trip
```

## Detailed User Guide

For a full walkthrough of all options (including screenshots, advanced settings, and round-trip examples), see the [docs/user-guide.md](docs/user-guide.md).

## License

[MIT](/LICENSE.md)

## Acknowledgments

Advanced text formatting features adapted from techniques used in 
[md2pptx](https://github.com/MartinPacker/md2pptx) by Martin Packer (MIT License).

Thanks to:
- https://asciiflow.com/#/ for their ascii wireframing tool
- https://plumberjack.blogspot.com/2019/11/a-qt-gui-for-logging.html for showing how to use Py's logging library with PySide6/Qt
- https://stackoverflow.com/questions/47666642/adding-an-hyperlink-in-msword-by-using-python-docx and https://github.com/python-openxml/python-docx/issues/384#issuecomment-294853130 for solutions to field-code hyperlink issues

## Known Limitations

See [docs/limitations.md](docs/limitations.md) for a detailed list of current limitations, unsupported features, and known workarounds.

## Troubleshooting

If conversion fails or the GUI won’t launch, see [docs/troubleshooting.md](docs/troubleshooting.md).


# Development & Contributing

To set up the project for development, see [Developer Guide](docs/dev-guide.md).
