
# ==DOCSTART==
# Purpose: Defines the UserConfig dataclass-- the single source of truth for user-overridable options.
# ==DOCEND==

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from enum import Enum
import os

# TODO: remove later
from docx2pptx_text.config import SCRIPT_DIR

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
    input_docx: Optional[str] = None # Use strings in the dataclass, convert to Path when you need to use them.
    input_pptx: Optional[str] = None

    output_folder: Optional[str] = None

    # ==> Templates I/O
    template_pptx: Optional[str] = None
    template_docx: Optional[str] = None


    # Processing
    chunk_type: ChunkType = ChunkType.HEADING_FLAT

    experimental_formatting_on: bool = True
    
    display_comments: bool = True    
    comments_sort_by_date: bool = True
    comments_keep_author_and_date: bool = True    

    display_footnotes: bool = True
    display_endnotes: bool = True
    
    preserve_docx_metadata_in_speaker_notes: bool = True


    # Behavior
    debug_mode: bool = True

    # Class methods
    def _resolve_path(self, raw: str) -> Path:
        """Expand ~ and ${VARS}; resolve relative to config_base_dir if present."""
        expanded = os.path.expandvars(raw)
        p = Path(expanded).expanduser()

        if p.is_absolute():
            return p.resolve()
        
        # For relative paths, resolve from repo root
        # TODO: (Later you'll use a proper base_dir from config)        
        base = SCRIPT_DIR
        return (base / p).resolve()
    
    def get_template_pptx_path(self) -> Path:
        """Get the template pptx path, with fallback to default."""
        if self.template_pptx:
            return self._resolve_path(self.template_pptx)
        
        # Default
        base = SCRIPT_DIR # TODO: replace with a proper base_dir
        return base / "resources" / "blank_template.pptx"