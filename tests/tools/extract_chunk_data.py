"""
Extract baseline JSON data from docx chunks for integration testing.

Usage:
    python tests/tools/extract_chunk_data.py

Uses internal Chunk_docx classes to test chunk-based processing.
"""

import json
import sys
from pathlib import Path

from extraction_utils import (  # type: ignore[import-not-found]
    filter_none_keep_false,
    rgb_to_hex,
    safe_pprint,
)

from manuscript2slides import io
from manuscript2slides.internals.define_config import ChunkType
from manuscript2slides.models import Chunk_docx
from manuscript2slides.processing.chunking import create_docx_chunks

# ============================================================================
# CONFIGURATION - Edit these to change input/output files
# ============================================================================
INPUT_DOCX = "tests/data/sample_doc.docx"
OUTPUT_JSON = "tests/baselines/docx_chunks_sample.json"
CHUNK_TYPE = (
    ChunkType.HEADING_FLAT
)  # Try: HEADING_FLAT, HEADING_NESTED, PAGE, PARAGRAPH
# ============================================================================


def main() -> None:
    """Load a test docx and explore its structure."""

    # Configuration
    test_docx_path = Path(INPUT_DOCX)
    chunk_type = CHUNK_TYPE

    if not test_docx_path.exists():
        print(f"Error: {test_docx_path} not found!", file=sys.stderr)
        print(
            f"Please check the INPUT_DOCX configuration at the top of this script.",
            file=sys.stderr,
        )
        sys.exit(1)

    # Load the docx
    print(f"Loading {test_docx_path}...")
    doc = io.load_and_validate_docx(test_docx_path)

    print(f"\nFound {len(doc.paragraphs)} paragraphs")
    print("\n=== Extracting fixture data ===")

    # Create chunks
    print(f"Creating chunks with chunk_type='{chunk_type.value}'...")
    chunks = create_docx_chunks(doc, chunk_type)
    print(f"Found {len(chunks)} chunks\n")

    data = []
    for chunk in chunks:
        chunk_data = {
            "original_sequence_number": chunk.original_sequence_number,
            "paragraphs": [],
        }

        for para in chunk.paragraphs:
            para_data: dict[str, object] = {
                "text": para.text,
                "style": para.style.name if para.style else None,
                "runs": [],
            }

            for run_idx, run in enumerate(para.runs):
                run_data = {
                    "run_number": run_idx,
                    "text": run.text,
                    "bold": run.font.bold,
                    "italic": run.font.italic,
                    "underline": run.font.underline,
                    "font_name": run.font.name,
                    "font_size": run.font.size.pt if run.font.size else None,
                }

                # Add color if present
                if run.font.color and run.font.color.rgb:
                    run_data["color_rgb"] = rgb_to_hex(run.font.color.rgb)

                # Extract experimental formatting from docx
                if run.font.highlight_color:
                    run_data["highlight_color"] = run.font.highlight_color.name

                if run.font.strike:
                    run_data["strike"] = run.font.strike

                if run.font.double_strike:
                    run_data["double_strike"] = run.font.double_strike

                if run.font.subscript:
                    run_data["subscript"] = run.font.subscript

                if run.font.superscript:
                    run_data["superscript"] = run.font.superscript

                if run.font.all_caps:
                    run_data["all_caps"] = run.font.all_caps

                if run.font.small_caps:
                    run_data["small_caps"] = run.font.small_caps

                # Filter out None values but keep False
                run_data = filter_none_keep_false(run_data)

                runs = para_data["runs"]
                assert isinstance(runs, list)
                runs.append(run_data)

            assert isinstance(chunk_data["paragraphs"], list)
            chunk_data["paragraphs"].append(para_data)

        # Add annotations if present
        if chunk.comments:
            chunk_data["comments"] = [
                {
                    "note_id": c.note_id,
                    "reference_text": c.reference_text,
                }
                for c in chunk.comments
            ]

        if chunk.footnotes:
            chunk_data["footnotes"] = [
                {
                    "note_id": f.note_id,
                    "text_body": f.text_body,
                    "hyperlinks": f.hyperlinks,
                    "reference_text": f.reference_text,
                }
                for f in chunk.footnotes
            ]

        if chunk.endnotes:
            chunk_data["endnotes"] = [
                {
                    "note_id": e.note_id,
                    "text_body": e.text_body,
                    "hyperlinks": e.hyperlinks,
                    "reference_text": e.reference_text,
                }
                for e in chunk.endnotes
            ]

        data.append(chunk_data)

    # Show what we got
    safe_pprint(data, width=120, item_type="chunks")

    # Save to JSON
    output_path = Path(OUTPUT_JSON)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {output_path}")
    except (OSError, PermissionError) as e:
        print(f"Error: Failed to write output file: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Fatal error: {e}", file=sys.stderr)
        sys.exit(1)
