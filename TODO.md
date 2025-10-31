## TODOs found while doing UI Exploration
- [x] refactor run_id: we need separate things considering run_id per-pipeline-run vs. session_id for per-UI session
- [x] Add logging to pipeline validation methods (validate_docx2pptx_pipeline_requirements, validate_pptx2docx_pipeline_requirements)
- [x] CLI: Add argparse support for --config flag
- [x] Change backend defaults to have most bools turned off so that default is empty speaker notes. Adjust round-trip to have them enabled.
- [ ] GUI: Wire up auto-save/load for preference persistence across sessions (call the save_toml/load_toml methods on UserConfig)
- [ ] Revisit the Demo speaker notes bool-setting code because I hard-coded the old way into cli.run() and it's... smelly....

## Major TODOs Ordered by priority/dependency
Epic: Investigate UI options, select one, implement it
- Build a simple UI with good enough UX that any non-tech-savvy writer can use it without friction

- [x] Finish toy GUI (Tkinter) to cover last few UI programming concepts
    - [x] tabbed view
    - [x] refactor to be component-based architecture
    - [-] ~~Custom events in tkinter~~
- [ ] Real Tkinter GUI (v1 real GUI)
    - Start real GUI work: Design
	- [x] Design it intentionally with wireframes
    - [x] Plan/describe the structure
        - What are all the features?
        - What are all the states?
        - How should it look?
	- [x] Outline the architecture conceptually, modularized, etc., without worrying about specific code syntax or framework
    - [x] v1 complete (Inheritance + Component mix)    
    - [x] Refactor Tabs to use MVP Pattern
    - [ ] (Optional) Refactor components to use MVP
    - [ ] Decide Tkinter is in a polished/good enough state to call done and use as an architectural reference for GUI v2, and move on.
        - [ ] Make sure you update dev-notes/ui_tree.txt and ui_wireframe.txt!!

- [ ] GUI Framework Exploration: PySide (Qt)
    - Think I want to fully swap to using PySide after finishing out Tkinter in a state I'm happy with.
        1) That'll help me reinforce/internalize what I learned during Tkinter build
        2) That'll give me experience with a more modern GUI framework
        3) Please, gods, let Qt solve the ugly theming and DPI issues of Tkinter
    - [ ] Work through some basic lessons on Qt/PySide, or directly adapt a component (LogViewer?) as a tutorial
    - RESIST THEMING/STYLING UNTIL YOU GET THE ARCHITECTURE DONE, JOJO!
    - [ ] Real GUI v2 build! Start fresh with a design and framework chosen intentionally
        - [ ] Progressively work through each bit of the old architecture to port to PySide


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