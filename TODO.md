## Major TODOs Ordered by priority/dependency

### Epic: Add tests & pytest
    - Test UI
    - Ensure we "catch" every possible UX-impacting raise/exception: for every "raise" possible to hit from the GUI, ensure we're popping message boxes and not crashing the app.

### Epic: Package/Distribution
    - Figure out how to package it into an installer for each non-mobile platform (Win 11, MacOS, Linux)
    - Make sure to update the resources/.. source inside of scaffold.py to use the packaged version
      - How do I support both the distro flow/file source + development using the source code resources/ folder? A const?

### Epic: Document the program thoroughly
    - [ ] Add "how to launch/run from vs code" into the docs/dev-guide.md 
    - for non-tech-savvy users
    - for future contributors
    - Don't forget to put the actual docstrings at the top of all the module files!

## Stretch Wishlist Features:
I'm not likely to get around to these, but I wanted to.
- Split the output pptx or docx into multiple output files based on slide or page count. Add default counts and allow user overrides for the default.
- Investigate if we can insert pptx sections safely enough (to allow for docx headings -> pptx sections, or other section-chunking); if not, investigate if/when we want to mimic the same type of behavior with "segue slides"
- Investigate how impossible non-local file input/output (OneDrive/SharePoint) would be; add to known limitations if not supportable.
- Investigate linking slides or sections-of-slides or file chunks back to their source "place" in the original docx (original file if possible, or a copy where we insert the anchor) (Maybe can just point to page of input file's absolute path)
- Add support to export to .md (1 file per chunk) to support docx -> zettelkasten workflows (Notion, Obsidian) (Consider supporting metadata -> YAML frontmatter)
- Add support to break chunks (of any type) at a word count threshold.
- Add support for importing .md and .txt; split by whitespaces or newline characters.
