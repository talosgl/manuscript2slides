"""TODO: add docstring"""
# ==DOCSTART==
# Purpose: Defines the UserConfig dataclass-- the single source of truth for user-overridable options.
# ==DOCEND==

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from enum import Enum
import os

# TODO: remove later
from docx2pptx_text.internals.constants import ROOT_DIR

# Which chunking method to use to divide the docx into slides. This enum lists the available choices:
class ChunkType(Enum):
    """Chunk Type Choices"""

    HEADING_NESTED = "heading_nested"
    HEADING_FLAT = "heading_flat"
    PARAGRAPH = "paragraph"
    PAGE = "page"


@ dataclass
class UserConfig:
    """All user-configurable settings for docx2pptx-text."""

    # Input/Output

    # Input file to process
    input_docx: Optional[str] = None # Use strings in the dataclass, convert to Path when you need to use them.
    input_pptx: Optional[str] = None 

    output_folder: Optional[str] = None # Desired output directory/folder to save in

    # ==> Templates I/O
    template_pptx: Optional[str] = None # The pptx file to use as the template for the slide deck
    template_docx: Optional[str] = None # The docx file to use as the template for the new docx


    # Processing
    chunk_type: ChunkType = ChunkType.HEADING_FLAT # Which chunking method to use to divide the docx into slides.

    experimental_formatting_on: bool = True
    
    display_comments: bool = True    
    comments_sort_by_date: bool = True
    comments_keep_author_and_date: bool = True    

    display_footnotes: bool = True
    display_endnotes: bool = True
    
    # We this way to leave speaker notes completely empty if the user really wants that, it's a valid use case.
    # Documentation and tooltips should make it clear that this means metadata loss for round-trip pipeline data.
    preserve_docx_metadata_in_speaker_notes: bool = True

    # Class methods
    def _resolve_path(self, raw: str) -> Path:
        """Expand ~ and ${VARS}; resolve relative to config_base_dir if present."""
        expanded = os.path.expandvars(raw)
        p = Path(expanded).expanduser()

        if p.is_absolute():
            return p.resolve()
        
        # For relative paths, resolve from repo root
        # TODO: (Later you'll use a proper base_dir from config)        
        base = ROOT_DIR
        return (base / p).resolve()
    
    # TODO: Consider: should templates even *be allowed* to be configureable by the user??
    def get_template_pptx_path(self) -> Path:
        """Get the docx2pptx template pptx path, with fallback to default."""
        if self.template_pptx:
            return self._resolve_path(self.template_pptx)
        
        # Default
        base = ROOT_DIR # TODO: replace with a proper base_dir
        return base / "resources" / "blank_template.pptx"
    
    
    def get_template_docx_path(self) -> Path:
        """Get the pptx2docx template docx path with fallback to a default."""
        if self.template_docx:
            return self._resolve_path(self.template_docx)
        
        # Default
        base = ROOT_DIR
        return base / "resources" / "docx_template.docx"


    # TODO: Consider collapsing these two input_file methods to match get_output_folder, rather than having different properties and methods per filetype.
    def get_input_docx_file(self) -> Path:
        """Get the docx2pptx input docx file or fall back to a dry run example file."""
        if self.input_docx:
            return self._resolve_path(self.input_docx)
        
        # Default/Dry Run
        base = ROOT_DIR
        return base / "resources" / "sample_doc.docx"


    def get_output_folder(self) -> Path:
        """Get the docx2pptx pipeline output pptx path, with fallback to default."""
        if self.output_folder:
            return self._resolve_path(self.output_folder)
        
        # Default
        base = ROOT_DIR
        return base / "output"

    
    def get_input_pptx_file(self) -> Path:
        """Get the pptx2docx input pptx file or fall back to a dry run example pptx."""
        if self.input_pptx:
            return self._resolve_path(self.input_pptx)
        
        # Default/Dry Run
        base = ROOT_DIR
        return base / "resources" / "sample_slides_output.pptx"
        
