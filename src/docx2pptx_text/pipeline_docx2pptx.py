"""main orchestrator"""

from docx2pptx_text.chunking import create_docx_chunks
from docx2pptx_text.annotations.extract import process_chunk_annotations
from docx2pptx_text.create_slides import slides_from_chunks
import sys
from pathlib import Path
from docx2pptx_text import io
from docx2pptx_text.internals.config.define_config import UserConfig

# TODO: replace docx_path throughout with cfg... and remove from signature
def run_docx2pptx_pipeline(cfg: UserConfig) -> None:
    """Orchestrates the docx2pptx pipeline."""
    user_path = cfg.get_input_docx_file()

    # Validate it's a real path of the correct type. If it's not, return the error.
    try:
        user_path_validated = io.validate_docx_path(user_path)
    except FileNotFoundError:
        print(f"Error: File not found: {user_path}")
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except PermissionError:
        print(f"I don't have permission to read that file ({user_path})!")
        sys.exit(1)

    # Load the docx file at that path.
    user_docx = io.load_and_validate_docx(user_path_validated)

    # Chunk the docx by ___
    chunks = create_docx_chunks(user_docx, cfg.chunk_type)

    if (
        cfg.display_comments or cfg.display_footnotes or cfg.display_endnotes
    ) or cfg.preserve_docx_metadata_in_speaker_notes:
        chunks = process_chunk_annotations(chunks, user_docx, cfg)

    # Create the presentation object from template
    try:
        output_prs = io.create_empty_slide_deck(cfg)
    except Exception as e:
        print(f"Could not load template file (may be corrupted): {e}")
        sys.exit(1)

    # Mutate the presentation object by adding slides
    slides_from_chunks(output_prs, chunks, cfg)

    # Save the presentation to an actual pptx on disk
    io.save_output(output_prs, cfg)
