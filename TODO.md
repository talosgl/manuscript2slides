## Minor TODOs (mostly independent)
- Finish logging implementation by replacing all the debug_print()s with actual logger calls.
- Rename the GH repo, finally! (Don't forget to nuke & redo the .venv/ and .vscode/ stuff, locally, probably...)

## Major TODOs Ordered by priority/dependency
Epic: Investigate UI options, select one, implement it
    - Build a simple UI with good enough UX that any non-tech-savvy writer can use it without friction

Epic: Package/Distribution
    - Figure out how to package it into an installer for each non-mobile platform (Win 11, MacOS, Linux)
    - Make sure to update the resources/.. source inside of scaffold.py to use the packaged version

Epic: Document the program thoroughly 
    - for non-tech-savvy users
    - for future contributors

## Public v1 Guidelines
- What does is "done enough for public github repo mean"? 
    - "When I'm comfortable having strangers use it without asking me questions."
    - Engineer-audience documentation
    - Log & error messages that tell users what went wrong and how to fix it
    - Code that doesn't crash on common edge cases.
- OK, but I think I really want it to *start* with UI in its first public version.

## Stretch Wishlist Features:
- Split the output pptx or docx into multiple output files based on slide or page count. Add default counts and allow user overrides for the default.
- Investigate if we can insert pptx sections safely enough (to allow for docx headings -> pptx sections, or other section-chunking); if not, investigate if/when we want to mimic the same type of behavior with "segue slides"
- Investigate how impossible non-local file input/output (OneDrive/SharePoint) would be; add to known limitations if not supportable.
- Investigate linking slides or sections-of-slides or file chunks back to their source "place" in the original docx (og file if possible, or a copy where we insert the anchor)
- Add support for importing .md and .txt; split by whitespaces or newline characters.
- Add support to break chunks (of any type) at a word count threshold.