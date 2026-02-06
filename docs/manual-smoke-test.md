# Manual Smoke Test

Basic checklist to verify core functionality, especially for binary output.

## Pre-Installation

- [ ] Download `manuscript2slides-windows.zip` from GitHub Releases
- [ ] Verify checksum matches the SHA256 hash in release notes
- [ ] Extract zip successfully (no errors)

## Launch

- [ ] `manuscript2slides.exe` launches without errors
- [ ] GUI opens and displays three tabs (DOCX → PPTX, PPTX → DOCX, Demo)

## DOCX → PPTX Tab

- [ ] **First launch folder scaffolding**
  - Click "Browse" for input file
  - Verify `~/Documents/manuscript2slides/` folders are created:
    - `input/`, `output/`, `templates/`, `logs/`, `configs/`
  - Select the sample DOCX from `input/`

- [ ] **Convert with all options enabled**
  - Enable: Experimental Formatting, Display Comments, Display Footnotes, Preserve Metadata
  - Click "Convert!"
  - Verify success message box appears
  - Check output PPTX:
    - Slides contain expected content
    - Speaker notes are populated with metadata
    - Formatting applied correctly

- [ ] **Convert with only experimental formatting**
  - Disable all options except Experimental Formatting
  - Click "Convert!"
  - Verify success message box appears
  - Check output PPTX:
    - Slides look correct
    - Speaker notes are empty (no metadata preserved)

## PPTX → DOCX Tab

- [ ] **Select input file**
  - Browse to a sample PPTX (use one generated in previous step or sample from `input/`)

- [ ] **Custom output folder**
  - Click "Browse" for output folder
  - Choose a different location (not default `output/`)

- [ ] **Convert and verify**
  - Click "Convert!"
  - Verify success message box appears
  - Confirm file was saved to the custom output folder (not default)
  - Open resulting DOCX and verify content looks correct

## Demo Tab

- [ ] **DOCX → PPTX Demo**
  - Click "Run DOCX → PPTX Demo"
  - Verify success and output file created

- [ ] **PPTX → DOCX Demo**
  - Click "Run PPTX → DOCX Demo"
  - Verify success and output file created

- [ ] **Round-trip Demo**
  - Click "Run Round-trip Demo"
  - Verify success and both output files created

- [ ] **Config-based Demo**
  - Load the sample `.toml` config from `~/Documents/manuscript2slides/configs/`
  - Run demo with config
  - Verify success and settings from config are applied

## Final Checks

- [ ] Check `~/Documents/manuscript2slides/logs/` - verify log files were created
- [ ] No crashes or unexpected errors throughout testing
- [ ] All output files open correctly in Word/PowerPoint

---
