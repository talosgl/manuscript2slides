
## Usage

### GUI Mode (Default)

Launch the application and you'll see three tabs:

**DOCX → PPTX Tab**

1. Select your Word document
2. Choose chunking method (paragraph, page, or heading-based)
3. Configure options (formatting, annotations, metadata)
4. Click "Convert!"

**PPTX → DOCX Tab**

1. Select your PowerPoint file
2. Click "Convert!"
3. Formatting and comments are restored if metadata is available

**Demo Tab**

- Try sample conversions without selecting files
- Test round-trip conversion (DOCX → PPTX → DOCX)

### CLI Mode

```bash
# Run demo with sample files
manuscript2slides-cli --demo-docx2pptx

# Convert a specific Word document
manuscript2slides-cli --input-docx my-manuscript.docx

# Convert PowerPoint to Word
manuscript2slides-cli --input-pptx presentation.pptx

# Use a configuration file
manuscript2slides-cli --config my-settings.toml

# Customize chunking strategy
manuscript2slides-cli --input-docx manuscript.docx --chunk-type heading_flat
```

## Save/Load Options
### Configuration Files

Save your preferences to avoid re-entering settings:

```toml
# my-config.toml
input_docx = "input/my-manuscript.docx"
chunk_type = "paragraph"
experimental_formatting_on = true
display_comments = true
display_footnotes = true
preserve_docx_metadata_in_speaker_notes = true
```

Then in the CLI, run:

```bash
manuscript2slides-cli --config my-config.toml
```

Or in the GUI, click "Advanced" under the input-file selector, and the "Save Config"/"Load Config" button. 

Sample configuration files are created automatically in `~/Documents/manuscript2slides/configs/` on first run.

### Persistent Preferences in GUI
In the GUI, we attempt to save/load selections across sessions automatically. You can clear saved selections from the Menu bar >  Preferences > Reset to Defaults, then relaunch the app.


## Chunking Strategies

- **Paragraph** (default): One slide per paragraph
- **Page**: One slide per page break
- **Heading (Flat)**: New slide at every heading, regardless of level
- **Heading (Nested)**: New slide only when reaching a "parent" heading level

All strategies create a new slide if a page break occurs mid-section.

## User Files Location

On first run, manuscript2slides creates a folder structure:

- **Windows**: `C:\Users\YourName\Documents\manuscript2slides\`
- **macOS**: `/Users/YourName/Documents/manuscript2slides/`
- **Linux**: `/home/yourname/Documents/manuscript2slides/`

```
manuscript2slides/
├── input/          # Sample files and staging area
├── output/         # Converted files (timestamped)
├── templates/      # Customizable DOCX/PPTX templates
├── logs/           # Log files for troubleshooting
└── configs/        # Saved configuration files
```


## Requirements

- **Python 3.10+** (if installing via pip)
- **Windows, macOS, or Linux**

### Linux-specific Dependencies

If you encounter Qt-related errors on Linux:

**Ubuntu/Debian:**

```bash
sudo apt install libxcb-xinerama0 libxcb-cursor0
```

**Fedora/RHEL:**

```bash
sudo dnf install xcb-util-cursor
```

## Getting Help

If you're having issues with the conversion process, [open an issue on GitHub](https://github.com/talosgl/manuscript2slides/issues) and include/attach:

- Your input file (if possible)
- The output file (if one was created)
- Log files from `~/Documents/manuscript2slides/logs/`
- Any manifest files from `~/Documents/manuscript2slides/manifests/`
- A description of what you expected vs. what happened

This information helps diagnose problems quickly.

(If we're buds IRL, you can just send me an email! But do try to attach these items if you can and it'll speed up the process.)