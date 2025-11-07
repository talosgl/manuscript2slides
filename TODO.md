## TODOs found while doing UI Exploration
- [ ] GUI: Wire up auto-save/load for preference persistence across sessions (call the save_toml/load_toml methods on UserConfig)
- [ ] Revisit the Demo speaker notes bool-setting code because I hard-coded the old way into cli.run() and it's... smelly....

## Major TODOs Ordered by priority/dependency
Feature: Add Provenance feature
    - Minimum: at the start of a pipeline call, dump the run_id/session_id + cfg (UserConfig) object's fields to log and/or file 

Epic: Add tests & pytest
    - Test config validation
    - Test path resolution across platforms
    - Test scaffolding (does it create folders? not overwrite files?)
    - Test the actual pipelines (docx->pptx, pptx->docx)
    - Test edge cases (empty docs, huge docs, corrupted files)
    - Maybe set up CI/CD? (GitHub Actions is free for public repos)

Epic: Package/Distribution
    - Figure out how to package it into an installer for each non-mobile platform (Win 11, MacOS, Linux)
    - Make sure to update the resources/.. source inside of scaffold.py to use the packaged version

Epic: Document the program thoroughly 
    - for non-tech-savvy users
    - for future contributors
    - Don't forget to put the actual docstrings at the top of all the module files!

## Public v1 Guidelines
- What does is "done enough for public github repo mean"? 
    - "When I'm comfortable having strangers use it without asking me questions."
    - Engineer-audience documentation
    - Log & error messages that tell users what went wrong and how to fix it
    - Code that doesn't crash on common edge cases.
- OK, but I think I really want it to *start* with UI in its first public version.

## Known Issues I'd like to investigate fixing:
- [ ] .docx Runs that are also Headings don't have their other formatting preserved when copied into the pptx _Run; just the fact it is a heading into the metadata. Perhaps we need to "get" the formatting details from the document's heading styles, rather than from the run's XML.
- [ ] Add feature to allow page ranges (Pipeline/backend + interface updates)

## Stretch Wishlist Features:
- Split the output pptx or docx into multiple output files based on slide or page count. Add default counts and allow user overrides for the default.
- Investigate if we can insert pptx sections safely enough (to allow for docx headings -> pptx sections, or other section-chunking); if not, investigate if/when we want to mimic the same type of behavior with "segue slides"
- Investigate how impossible non-local file input/output (OneDrive/SharePoint) would be; add to known limitations if not supportable.
- Investigate linking slides or sections-of-slides or file chunks back to their source "place" in the original docx (og file if possible, or a copy where we insert the anchor)
- Add support for importing .md and .txt; split by whitespaces or newline characters.
- Add support to export to .md (1 file per chunk)
- Add support to break chunks (of any type) at a word count threshold.