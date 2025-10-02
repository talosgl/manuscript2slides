# Overarching TODOs
Must-Implement v0 Features:
- Change consts configuration to use a class or similar
- Rearchitect to be multi-file

v1 features I'd like:
- Add a feature to split the output pptx or docx into multiple files based on slide or page count. Add default counts and allow user overrides for the default.
- Investigate if we can insert pptx sections safely enough (to allow for docx headings -> pptx sections, or other section-chunking)
    - If not, investigate if/when we want to mimic the same type of behavior with "segue slides"
- Create Documents/docx2pptx/input/output/resources structure
    - Copies sample files from app resources to user folders
    - Cleanup mode for debug runs    
- Add an actual logger
- Investigate how impossible non-local file input/output (OneDrive/SharePoint) would be; add to known limitations if not supportable.

Public v1
- What does is "done enough for public github repo mean"? 
    - "When I'm comfortable having strangers use it without asking me questions."
    - Engineer-audience documentation
    - Log & error messages that tell users what went wrong and how to fix it
    - Code that doesn't crash on common edge cases.

Public v2: UI
    - Build a simple UI with good enough UX that any non-tech-savvy writer can use it without friction

Public v3: Package/Distribution
    - Figure out how to package it into an installer for each non-mobile platform (Win 11, MacOS, Linux)

Stretch Wishlist Features:
-   Investigate linking slides or sections-of-slides or file chunks back to their source "place" in the original docx (og file if possible, or a copy where we insert the anchor)
-   Add support for importing .md and .txt; split by whitespaces or newline characters.
-   Add support to break chunks (of any type) at a word count threshold.