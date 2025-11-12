## Major TODOs Ordered by priority/dependency
Epic: Bucket o' Fixes & Features pre-v1:

- GUI-only Feature:
  - [ ]  Preference Persistence: Wire up auto-save/load for preference persistence across sessions; use QtSettings to let users decide if preferences / app state should persist across sessions or be cleared every time.    
    - Default this behavior to remember/persist: "Users generally prefer software that restores the state they were last in. It minimizes friction and assumes a typical workflow where users iterate on a small set of inputs rather than starting from a blank slate every single time. The vast majority of well-known software (browsers, IDEs, office suites) remember window size, recent files, and input fields. Your users will expect this behavior."
    - Use QMenuBar + QSettings to offer a preference persistence option for the user; place the access to this feature in a standard Menubar location (`Edit` -> `Preferences`), not a Toolbar.
    - Additionally, provide a way for the users to quickly clear options/reset fields to defaults in the main UI.

- Backend-only Features:
  - [x] .docx Runs that are also Headings don't have their other formatting preserved when copied into the pptx _Run; just the fact it is a heading into the metadata. Perhaps we need to "get" the formatting details from the document's heading styles, rather than from the run's XML.
  - [ ] Incorporate Pydantic into the project for better automatic validation, type-checking, etc.
    - Pylance strict reports some issues with over-validation define_config. One suggested fix/path forward is to be using Pydantic instead of doing manual checks in my code. Additionally, it's a low-impact library and will be useful to learn in general because it is an industry-standard library.
- [ ] Make DEBUG_MODE not just a source code const
  
- GUI/CLI + Backend Features: 
  - [ ] Add feature to allow page ranges

### Epic: Add tests & pytest
    - Test config validation
    - Test path resolution across platforms
    - Test scaffolding (does it create folders? not overwrite files?)
    - Test the actual pipelines (docx->pptx, pptx->docx)
    - Test edge cases (empty docs, huge docs, corrupted files)
    - Maybe set up CI/CD? (GitHub Actions is free for public repos)

### Epic: Package/Distribution
    - Figure out how to package it into an installer for each non-mobile platform (Win 11, MacOS, Linux)
    - Make sure to update the resources/.. source inside of scaffold.py to use the packaged version

### Epic: Document the program thoroughly
    - [ ] Add "how to launch/run from vs code" into the docs/dev-guide.md 
    - for non-tech-savvy users
    - for future contributors
    - Don't forget to put the actual docstrings at the top of all the module files!

## Stretch Wishlist Features:
- Split the output pptx or docx into multiple output files based on slide or page count. Add default counts and allow user overrides for the default.
- Investigate if we can insert pptx sections safely enough (to allow for docx headings -> pptx sections, or other section-chunking); if not, investigate if/when we want to mimic the same type of behavior with "segue slides"
- Investigate how impossible non-local file input/output (OneDrive/SharePoint) would be; add to known limitations if not supportable.
- Investigate linking slides or sections-of-slides or file chunks back to their source "place" in the original docx (og file if possible, or a copy where we insert the anchor)
- Add support to export to .md (1 file per chunk) to support docx -> zettelkasten workflows (Notion, Obsidian) (Consider supporting metadata -> YAML frontmatter)
- Add support to break chunks (of any type) at a word count threshold.
- Add support for importing .md and .txt; split by whitespaces or newline characters.
- Investigate Paragraph(docx).style.paragraph_format.keep_together / keep_with_next for chunking vs headings/page break