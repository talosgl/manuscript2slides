"""
Test Fixture Generation Tools

This package contains scripts for generating JSON baseline fixtures from docx/pptx
files for integration testing. These are internal development tools, not part of
the public API.

## Purpose

These scripts extract comprehensive data (formatting, annotations, structure) from
Office documents into JSON files that serve as test baselines. The baselines are
committed to git and used by integration tests to verify the manuscript2slides
pipeline produces expected output.

## Available Scripts

### extract_docx_data.py
Extracts paragraph, run, and annotation data from a .docx file.

**What it captures:**
- Paragraph-level formatting (from styles)
- Run-level formatting (bold, italic, fonts, colors, etc.)
- Experimental formatting (highlight, strikethrough, caps, sub/superscript)
- Annotations (comments, footnotes, endnotes) with their reference locations

**Usage:**
```bash
python tests/tools/extract_docx_data.py
```

Edit the constants at the top of the file to change input/output paths:
- `INPUT_DOCX`: Path to source .docx file
- `OUTPUT_JSON`: Path to output baseline JSON

### extract_pptx_data.py
Extracts slide, shape, and speaker notes data from a .pptx file.

**What it captures:**
- Slide structure with shapes and text frames
- Paragraph and run-level formatting
- Speaker notes with parsed metadata (annotations, headings, etc.)
- Experimental XML formatting not exposed by python-pptx

**Usage:**
```bash
python tests/tools/extract_pptx_data.py
```

Edit the constants at the top of the file to change input/output paths:
- `INPUT_PPTX`: Path to source .pptx file
- `OUTPUT_JSON`: Path to output baseline JSON

### extract_chunk_data.py
Extracts chunk-based document data using internal Chunk_docx classes.

**What it captures:**
- Document chunks based on configurable chunking strategy
- Paragraph and run data within each chunk
- Annotations associated with each chunk

**Usage:**
```bash
python tests/tools/extract_chunk_data.py
```

Edit the constants at the top of the file to change input/output paths:
- `INPUT_DOCX`: Path to source .docx file
- `OUTPUT_JSON`: Path to output baseline JSON
- `CHUNK_TYPE`: Chunking strategy (HEADING_FLAT, HEADING_NESTED, PAGE, PARAGRAPH)

## Typical Workflow

1. **Initial fixture generation:**
   - Run the appropriate extraction script(s) for your test documents
   - Review the generated JSON to ensure it captures what you need
   - Commit the baseline JSON files to git

2. **Writing tests:**
   - Use the committed JSON baselines in your integration tests
   - Tests compare actual output against these known-good baselines

3. **Updating fixtures:**
   - When adding features or fixing bugs, you may need to regenerate baselines
   - Edit the source .docx/.pptx files as needed
   - Re-run the extraction scripts
   - Review the git diff to verify changes are expected
   - Commit the updated baselines

## Implementation Notes

- These scripts use manuscript2slides internal functions for extraction
- They are NOT independent of the codebase being tested
- Baselines are generated manually and reviewed for correctness
- Tests remain independent by comparing against committed JSON (not regenerating on each test run)
- The extraction_utils.py module provides shared helpers to reduce duplication

## Private Module

This package is for internal development use only. Nothing should be imported
from this module by other code. If you need these utilities elsewhere, consider
whether they belong in the main codebase instead.
"""

# Empty __all__ signals this is a private module - nothing should be imported
__all__ = []
