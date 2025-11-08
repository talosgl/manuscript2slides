
# manuscript2slides

Convert Microsoft Word documents to PowerPoint slides and back again.

## Features

- **Multiple chunking strategies**: By paragraph, page, or heading (flat/nested)
- **Formatting preservation**: Bold, italics, colors, highlights, strikethrough, super/subscript, and more
- **Annotation support**: Comments, footnotes, and endnotes can be copied to slide speaker notes
- **Round-trip capability**: Convert DOCX → PPTX → DOCX with optional metadata preservation
- **Cross-platform**: Works on Windows, macOS, and Linux
- **Both GUI and CLI**: Use whichever fits your workflow

## Quick Start Guide

### Download Executable (Recommended)

1. Download the latest release for your platform from the [latest release](https://github.com/talosgl/manuscript2slides/releases/latest) page.

Just download, extract (if needed), and run. No installation required.

2. Open `manuscript2slides` (no install needed).
3. Choose a tab:
   - **DOCX → PPTX:** Convert your manuscript into slides.
   - **PPTX → DOCX:** Turn a slide deck back into text.
   - **DEMO:** Try a sample conversion.

4. Pick your file(s), adjust options (chunking, formatting, annotations), and click **Convert!**

### (Advanced) CLI Guide

You can install manuscript2slides via the executables above, or with pip:
```bash
pip install manuscript2slides
```

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
[docs/user-guide.md](docs/user-guide.md)

## License

[MIT](/LICENSE.md)

## Acknowledgments

Advanced text formatting features adapted from techniques used in 
[md2pptx](https://github.com/MartinPacker/md2pptx) by Martin Packer (MIT License).

## Known Limitations

See [docs/limitations.md](docs/limitations.md) for a detailed list of current limitations, unsupported features, and known workarounds.

## Troubleshooting

If conversion fails or the GUI won’t launch, see [docs/troubleshooting.md](docs/troubleshooting.md).


# Development & Contributing

To set up the project for development, install from source, or run tests, see [docs/dev-guide.md]