# docx2slides-py README

## What is doc2slides-py?
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

🛠️ TODO: Build out outline below
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