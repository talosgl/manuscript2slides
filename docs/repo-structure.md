# Repository Structure

This file provides an overview of the manuscript2slides codebase structure with annotations explaining each file's purpose.

Last updated v0.1.5

---
```
manuscript2slides/
├──BUILD.md                                       # Quick reference for building platform binaries with Nuitka
├──MANIFEST.in                                    # Specifies additional files to include in Python source distributions
├──README.md                                      # Project overview, features, installation, and quick start guide
├──build.py                                       # Cross-platform build script for creating binaries using Nuitka
├──nuitka-package.config.yaml                     # Nuitka build patches for python-pptx/docx path resolution on macOS
├──mypy.ini                                       # Type checking configuration for mypy static analyzer
├──pyproject.toml                                 # Python project metadata, dependencies, and build system configuration
├──pytest.ini                                     # pytest test runner configuration and settings
├──requirements-binary-build.txt                  # Nuitka and other dependencies needed for building standalone binaries
├──docs/                                          # User-facing documentation                                    
|  ├──building.md                                 # Instructions for building platform binaries with Nuitka
|  ├──code-structure.md                           # Overview of codebase structure with file annotations
|  ├──dev-guide.md                                # Developer guide for installation, testing, and contribution
|  ├──limitations.md                              # Known limitations and design constraints
|  ├──manual-smoke-test.md                        # Manual smoke test checklist for binary releases
|  ├──releasing.md                                # Release process for PyPI packages and GitHub releases
|  ├──troubleshooting.md                          # Troubleshooting guide for common issues
|  ├──user-guide.md                               # User guide for GUI and CLI usage
|  └──dev-process-archive/...                     # Archived process artifacts from during dev
├──src/                                           # Top-level directory for main codebase source code; follows setuptools' suggested src-layout (https://setuptools.pypa.io/en/latest/userguide/package_discovery.html#src-layout)
|  └──manuscript2slides/                           
|     ├──__init__.py                              # Package marker; also handles version extraction from pyproject.toml
|     ├──__main__.py                              # Entry point for manuscript2slides desktop application.
|     ├──cli.py                                   # CLI Interface Logic (argparse etc)
|     ├──file_io.py                               # File I/O operations for docx and pptx files.
|     ├──gui.py                                   # Main GUI entry point for manuscript2slides
|     ├──models.py                                # Data models for document chunks and annotations.
|     ├──orchestrator.py                          # Route program flow to the appropriate pipeline based on user-indicated direction.
|     ├──startup.py                               # Startup logic needed by both CLI and GUI interfaces before anything else happens.
|     ├──templates.py                             # Load docx and pptx templates from disk, validate shape, and create in-memory python objects from them.
|     ├──utils.py                                 # Utilities for use across the entire program.
|     ├──annotations/
|     |  ├──__init__.py
|     |  ├──apply_to_slides.py                    # Add annotations we pulled from the docx to PowerPoint slide notes.
|     |  ├──extract.py                            # Extract annotations from Word documents.
|     |  └──restore_from_slides.py                # Restore annotations from slide metadata.
|     ├──archive/
|     |  ├──__init__.py
|     |  └──ui_tk.py                              # ARCHIVED: Tkinter and ttk GUI interface entry point. I'm just sentimental because it was my first GUI attempt.
|     ├──internals/
|     |  ├──__init__.py                           # Internal utilities and configuration for manuscript2slides.
|     |  ├──constants.py                          # Application-wide constants and configuration values.
|     |  ├──define_config.py                      # User configuration dataclass and validation.
|     |  ├──logger.py                             # Basic logging setup; creates console and file handlers with session_id in every log line.
|     |  ├──manifest.py                           # Track and record metadata for pipeline runs.
|     |  ├──paths.py                              # Cross-platform path resolution for user directories.
|     |  ├──run_context.py                        # Process-global execution context management.
|     |  └──scaffold.py                           # User directory structure creation and initialization.
|     ├──pipelines/
|     |  ├──__init__.py
|     |  ├──docx2pptx.py                          # Word to PowerPoint conversion pipeline.
|     |  └──pptx2docx.py                          # PowerPoint to Word conversion pipeline.
|     ├──processing/
|     |  ├──__init__.py
|     |  ├──chunking.py                           # Create docx2pptx chunks including by paragraph, page, heading (flat), and heading (nested).
|     |  ├──create_slides.py                      # Take chunks we built from the input docx file and turn them into slide body content.
|     |  ├──docx_xml.py                           # Docx XML parsing utilities for extracting data exposed by existing interop libraries.
|     |  ├──formatting.py                         # Formatting functions for both pipelines.
|     |  ├──populate_docx.py                      # Process slides from a presentation and copy their content into a Word document.
|     |  └──run_processing.py                     # Processes inner-paragraph contents (runs, hyperlinks) for both pipeline directions.
|     └──resources/
|        ├──_resources.md                         # `resources/` contains templates and sample files get copied into the user's ~/Documents/manuscript2slides/ directories.
|        ├──docx_template.docx                    # Default Word template for PPTX→DOCX conversions
|        ├──pptx_template.pptx                    # Default PowerPoint template for DOCX→PPTX conversions
|        ├──sample_doc.docx                       # Sample document for testing and demos
|        └──scaffold_README.md                    # README copied to user's ~/Documents/manuscript2slides/
└──tests/
   ├──__init__.py                                 # Test suite for manuscript2slides.
   ├──conftest.py                                 # Shared fixtures
   ├──helpers.py                                  # Shared test helper functions.
   ├──test_cli.py                                 # Tests for CLI argument parsing and config building.
   ├──test_file_io.py                             # Test I/O functions
   ├──test_gui.py                                 # Baseline Tests for the GUI
   ├──test_integration.py                         # Wide scoped tests that touch many parts of the pipeline and program to ensure no catastrophic failures during golden path runs.
   ├──test_main.py                                # Tests for application entry point routing when calling `python manuscript2slides`
   ├──test_smoke.py                               # Smoke tests to ensure basic functionality works.
   ├──test_startup.py                             # Tests for application startup.
   ├──test_templates.py                           # Tests to ensure we can create blank slide deck and blank document from our standard templates.
   ├──test_utils.py                               # Tests for utility functions.
   ├──annotations/
   |  ├──test_footnote_regression.py              # Regression test for the '1. 1' footnote double numbering issue.
   |  └──test_restore_from_slides.py              # Tests for restore_from_slides module - JSON parsing, string manipulation, and range merging.
   ├──internals/
   |  ├──test_define_config.py                    # Tests for UserConfig class definition file and related items.
   |  ├──test_manifest.py                         # Test the manifest system.
   |  └──test_scaffold.py                         # Tests to ensure the manuscript2slides directory structure gets created properly under the users' Documents.
   ├──processing/
   |  ├──test_chunking.py                         # Tests for all chunking strategies as well as helpers that perform heading detection.
   |  └──test_formatting.py                       # Tests for formatting in both pipeline directions.
   └──data/..., and baselines/..., tools/...      # Test fixture resources
```
