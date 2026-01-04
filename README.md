
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

### Installation

**Note:** Standalone executables for Mac and Windows (no Python required) are planned for a future release.

For now, you can install via pip (requires Python 3.10+ to already be installed):

```bash
# Recommended: Create a virtual environment first
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Then install
pip install manuscript2slides
```



### Using the GUI (recommended for most users)

```bash
# Launch the graphical interface
manuscript2slides
```

Then:
1. Choose a tab:
   - **DOCX → PPTX:** Convert your manuscript into slides
   - **PPTX → DOCX:** Turn a slide deck back into text
   - **DEMO:** Try a sample conversion
2. Pick your file(s), adjust options (chunking, formatting, annotations), and click **Convert!**

### Using the Command Line

```bash
# Convert a Word document
manuscript2slides-cli --input-docx my-manuscript.docx

# Reverse conversion
manuscript2slides-cli --input-pptx presentation.pptx --direction pptx2docx

# See a demo dry run with sample files
manuscript2slides-cli --demo-round-trip
```

## Detailed User Guide

For a full walkthrough of all options (including screenshots, advanced settings, and round-trip examples), see the [User Guide](https://github.com/talosgl/manuscript2slides/blob/main/docs/user-guide.md).

## License

[MIT](https://github.com/talosgl/manuscript2slides/blob/main/LICENSE)

## Acknowledgments

Advanced text formatting features adapted from techniques used in 
[md2pptx](https://github.com/MartinPacker/md2pptx) by Martin Packer (MIT License).

Thanks to:
- [ASCIIFlow](https://asciiflow.com/#/) for their ascii wireframing tool
- Blog post: [A Qt GUI for Logging](https://plumberjack.blogspot.com/2019/11/a-qt-gui-for-logging.html) for showing how to use Py's logging library with PySide6/Qt
- [StackOverflow answer](https://stackoverflow.com/questions/47666642/adding-an-hyperlink-in-msword-by-using-python-docx) and [GitHub discussion](https://github.com/python-openxml/python-docx/issues/384#issuecomment-294853130) for guidance on advanced techniques to add hyperlinks to docx runs

For full licensing details, see [THIRD_PARTY_LICENSES.md](https://github.com/talosgl/manuscript2slides/blob/main/THIRD_PARTY_LICENSES.md).

## Known Limitations

See [Known Limitations](https://github.com/talosgl/manuscript2slides/blob/main/docs/limitations.md) for a detailed list of current limitations, unsupported features, and known workarounds.

## Troubleshooting

If conversion fails or the GUI won't launch, see [Troubleshooting](https://github.com/talosgl/manuscript2slides/blob/main/docs/troubleshooting.md).


# Development & Contributing

To set up the project for development, see [Developer Guide](https://github.com/talosgl/manuscript2slides/blob/main/docs/dev-guide.md).
