## What is doc2pptx_text?
A Python script that converts a Microsoft Word .docx manuscript file into a set of basic PowerPoint .pptx text-frame slides.

The script requires a bit of setup, but is designed to be easy-to-use for non-devs; all the setup needed is covered below.

## Features
A manuscript is chunked into slides by paragraphs by default, but can also be chunked by pages, or by headings (nested or flat). Optional formatting preservation keeps bold, italics, and font colors.

A blank pptx template is provided that's tested to work well with the script. You can customize this with your preferred fonts, etc., or provide a different file. (However, know that significant alterations may require you to update the code yourself.)

## Quickstart
If you're familiar with command line and Python already, you can see a dry run by cloning the repo locally and running:
```bash
python docx2slides.py  # Uses sample files from resources/ and saves output to output/
```


## How to set up doc2slides-py
The idea with the sections below is that they build on one another. If you're already familiar with programming and open source workflows, you can skip the early sections and pick up where it makes sense. If you're completely new to all this, just start at the beginning to get fully set up.

### Get VS Code (IDE), get the program source code, get comfortable in the UI
*Audience: You're tech savvy, but new to programming, command-line, and/or git workflows*

- Get an IDE (we'll use VS Code in the guide)
- Clone or download the code
- Open it in your IDE
- Understand what the console/shell is

### Get Python installed & learn about venv
*Audience: Comfortable with programming but new to Python*

- Install Python
- Install Python extension for VS Code
- Learn/understand what Python virtual environments are

### Run docx2slides demo
*Audience: Ready to run/dev in Python and familiar with virtual environments*

- Set up a venv for this project
- pip install dependencies: `pip install -e .`
- Run docx2slides demo with sample doc

### Customize and run docx2slides for your use case
*Audience: You've got the demo working and are ready to customize*

- Customize the script for your use case using the constants at the top
- Rinse & repeat!


## Example Use Case: Manuscript Review Workflow:
Convert your Word manuscript to slides, then use PowerPoint's Notes View to review and annotate each section. Print or export as PDF for offline review. This creates a unique workflow for iterative manuscript editing.

## Limitations

- Comments, footnotes, and endnotes are not supported
- Complex Word formatting may not transfer perfectly
- Only text content is tested; images, tables, charts, etc. probably don't work
- Reverse conversion (PowerPoint to Word) is not supported. I plan to try implementing this and other features later in a C# version of the program.

## Support and Ongoing Development
This is a personal project-- no guarantees on support or updates.

Still, you're welcome to submit issues if you get stuck or have feature suggestions; although I cannot commit to responding to them, someone else might be able to help.

Small bug fixes and feature additions are welcome as PRs, but architectural changes should probably be forked.

(If I sent you this script because we're writing buds, feel free reach out to me directly. Sorry for all the tech setup!)


## Acknowledgments
Advanced text formatting features adapted from techniques used in 
[md2pptx](https://github.com/MartinPacker/md2pptx) by Martin Packer (MIT License).



====

## Previous docstring atop single-file program version
Convert Microsoft Word documents to PowerPoint presentations.

This tool processes .docx files and converts them into .pptx slide decks by chunking
the document content based on various strategies (paragraphs, headings, or page breaks).
Text formatting like bold, italics, and colors can optionally be preserved.

The main workflow:
1. Load a .docx file using python-docx
2. Chunk the content based on the selected strategy
3. Create slides from chunks using a PowerPoint template
4. Save the resulting .pptx file

Supported chunking methods:
- paragraph: Each paragraph becomes a slide
- page: New slide for each page break
- heading_flat: New slide for each heading (any level)
- heading_nested: New slide based on heading hierarchy

Example:
    python manuscript2slides.py

    (Configure INPUT_DOCX_FILE and other constants before running)


====

## Known Issues & Limitations
-   We only support text content. No images, tables, etc., are copied between the formats, and we do not have plans 
    to support these in future.

- "Sections" in both docx and pptx are not supported. TODO, leafy: investigate

-   We do not support .doc or .ppt, only .docx. If you have a .doc file, convert it to .docx using Word, Google Docs, 
    or LibreOffice before processing.

-   We do not support .ppt, only .pptx.

-   Auto-fit text resizing in slides doesn't work. PowerPoint only applies auto-fit sizing when opened in the UI. 
    You can get around this manually with these steps:
        1. Open up the output presentation in PowerPoint Desktop > View > Slide Master
        2. Select the text frame object, right-click > Format Shape
        3. Click the Size & Properties icon {TODO, doc: ADD SCREENCAPS}
        4. Click Text Box to see the options
        5. Toggle "Do not Autofit" and then back to "Shrink Text on overflow"
        6. Close Master View
        7. Now all the slides should have their text properly resized.

-   Field code hyperlinks not supported - Some hyperlinks (like the sample_doc.docx's first "Where are Data?" link) 
    are stored using Word's field code format and display as plain text instead of clickable links. The exact 
    conditions that cause this format are unclear, but it may occur with hyperlinks in headings or certain copy/paste 
    scenarios. We think most normal hyperlinks will work fine. We try to report when we detect these are present but cannot
    reliably copy them as text into the body.

-   ANNOTATIONS LIMITATIONS
    -   We collapse all comments, footnotes, and endnotes into a slide's speaker notes. PowerPoint itself doesn't 
        support real footnotes or endnotes at all. It does have a comments functionality, but the library used here 
        (python-pptx) doesn't support adding comments to slides yet. 

    -   Note that inline reference numbers (1, 2, 3, etc.) from the docx body are not preserved in the slide text - 
        only the annotation content appears in speaker notes.

    -   You can choose to preserve some comment metadata (author, timestamps) in plain text, but not threading.

-   REVERSE FLOW LIMITATIONS
    -   The reverse flow (pptx2docx-text) is significantly less robust. Your original input document to the manuscript2slides flow, and 
        the output document from a follow-up pptx2docx-text flow will not look the same. Expect to lose images, tables, footnotes, 
        endnotes, and fancy formatting. We attempt to preserve headings (text-matching based). Comments should be restored, but their 
        anchor positioning may be altered slightly.

    -   There will always be a blank line at the start of the reverse-pipeline document. When creating a new document with python-docx 
        using Document(), it inherently includes a single empty paragraph at the start. This behavior is mandated by the Open XML 
        .docx standard, which requires at least one paragraph within the w:body element in the document's XML structure.
