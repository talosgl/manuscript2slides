## TODOs found while doing UI Exploration
- [ ] refactor run_id: we need separate things considering run_id per-pipeline-run vs. session_id for per-UI session
- [x] refactor UserConfig() and move dry run functionality into a config that can be loaded from a button in the UI / a class method from the CLI, rather than being "magically" populating as the default values
    - [x] refactor to use toml instead of yaml for configs
    - [x] add save/load config functionality
    - [x] add with_defaults() class method for CLI to use
    - [x] update validation methods to not assume defaults (e.g., change input_docx to fail on no input file rather than auto-populate)
- [x] Whatever backend work is needed to support "preference persistence" (auto-save user UI config selections on close?)
- [ ] Should `__main__.py` have the log = setuplogger() thing at the top, like all the other files ... just in case?
- [ ] Add logging to pipeline validation methods (validate_docx2pptx_pipeline_requirements, validate_pptx2docx_pipeline_requirements)
- [ ] CLI: Add argparse support for --config flag
- [ ] GUI: Wire up auto-save/load for preference persistence across sessions (simple, just call the save_toml/load_toml methods we built)

## Major TODOs Ordered by priority/dependency
Epic: Investigate UI options, select one, implement it
- Build a simple UI with good enough UX that any non-tech-savvy writer can use it without friction

- [ ] Learn basic UI programming concepts by doing Tkinter mini-tutorials in the context of manuscript2slides
    - [x] clean separation between UI frontend from backend/config (had this already because of backend architectural reasons, but now understand how it matters in context of UI)
    - [x] Event-driven programming
    - [x] UI state management
    - [x] Layout systems (grid)
    - [x] File dialogs
    - [x] Dynamic UI updates
    - [ ] Error handling in UI context
    - [x] "Persistent state" pattern - implemented with TOML save/load
    - [ ] Explicit actions > implicit "magic" when it comes to UI apps. (If we want to give users a dry run feature, we should have them click a button that makes it clear they're invoking that. Not just auto-call it, like we might with a cli command.)

- [ ] finish Tkinter prototype/experiment - finish a complete, functional prototype.
    - [ ] Log viewer
    - [ ] Threading (so UI doesn't freeze)
    - [ ] Success/error message boxes
    - Stretch:
    - [ ] Progress indicator (spinning wheel or progress bar)
    - [ ] Show output location on success
    - Stretch even more:
    - [ ] Try some Tkinter templates to see if we can make it look modern (`from tkinter import ttk`) # ttk widgets look more modern

- [ ] DESIGN BREAK
    - Take a break from UI coding
    - Design what you actually want:
    - [x] Sketch it on paper / Wireframe what you want
    - [ ] Plan/describe the structure
        - What are all the features?
        - What are all the states?
        - How should it look?
    - [ ] Evaluate frameworks:
        - [ ] try PyQt
        - [ ] try Eel

- [ ] Start fresh with a design and framework chosen intentionally

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


## Stretch Wishlist Features:
- Split the output pptx or docx into multiple output files based on slide or page count. Add default counts and allow user overrides for the default.
- Investigate if we can insert pptx sections safely enough (to allow for docx headings -> pptx sections, or other section-chunking); if not, investigate if/when we want to mimic the same type of behavior with "segue slides"
- Investigate how impossible non-local file input/output (OneDrive/SharePoint) would be; add to known limitations if not supportable.
- Investigate linking slides or sections-of-slides or file chunks back to their source "place" in the original docx (og file if possible, or a copy where we insert the anchor)
- Add support for importing .md and .txt; split by whitespaces or newline characters.
- Add support to break chunks (of any type) at a word count threshold.