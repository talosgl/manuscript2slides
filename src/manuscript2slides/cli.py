"""CLI Interface Logic (argparse etc)"""

import argparse
from argparse import ArgumentError, ArgumentTypeError
import sys
from dataclasses import fields
from pathlib import Path

from manuscript2slides.utils import setup_console_encoding
from manuscript2slides.internals.config.define_config import (
    UserConfig,
    ChunkType,
    PipelineDirection,
)

# we'll need this later to replace run_roundtrip_test
from manuscript2slides.orchestrator import run_pipeline

import logging

log = logging.getLogger("manuscript2slides")


def run() -> None:
    """Run CLI interface. Assumes startup.initialize_application() was already called."""

    # Parse command line arguments
    args = parse_args()

    # Build config from args (with proper prioritization CLI args > config file > defaults)
    cfg = build_config_from_args(args)

    if args.demo_round_trip:
        log.info("Running round-trip test with sample files.")

        from manuscript2slides.orchestrator import run_roundtrip_test

        run_roundtrip_test(cfg)
    else:
        run_pipeline(cfg)


def parse_args() -> argparse.Namespace:
    """
    Parse command line arguments.

    Returns argparse.Namespace with all the UserConfig fields as attributes.
    Validates that all config fields hav corresponding CLI arguments.
    """
    parser = argparse.ArgumentParser(
        prog="manuscript2slides",
        description="Convert text content from Word docx to PowerPoint pptx slides and vice versa",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Use config file
  manuscript2slides --config path/to/my_settings.toml

  # See a demo run with sample files
  manuscript2slides --demo-docx2pptx
  
  # Quick conversion of a real file with defaults
  manuscript2slides --input-docx manuscript.docx
  
  # Override config file settings
  manuscript2slides --config settings.toml --direction pptx2docx
        """,
    )

    # Args for running dry runs
    parser.add_argument(
        "--demo-docx2pptx",
        action="store_true",
        dest="demo_docx2pptx",
        help=(
            "Run a demonstration conversion from a sample Word document to PowerPoint slides. "
            "Ignores other CLI options and uses built-in example files."
        ),
    )
    parser.add_argument(
        "--demo-pptx2docx",
        action="store_true",
        dest="demo_pptx2docx",
        help="Run a demonstration conversion from sample PowerPoint slides to a Word document. "
        "Ignores other CLI options and uses built-in example files.",
    )

    # Arg for running dry run
    parser.add_argument(
        "--demo-round-trip",
        action="store_true",
        dest="demo_round_trip",
        help=(
            "Run both conversions (docx -> pptx and pptx -> docx) using sample files, back to back. "
            "Useful for testing end-to-end behavior. Ignores other CLI options."
        ),
    )

    # Config file (special - loads other values)
    parser.add_argument(
        "--config",
        type=str,
        metavar="PATH",  # This prepends the help text and tells the user this is intended to be a file path
        help="Path to TOML configuration file for a pipeline run. See example in ~/Documents/manuscript2slides/configs/sample_config.toml after at least 1 run",
    )

    # Input/Output files
    parser.add_argument(
        "--input-docx",
        type=str,
        dest="input_docx",
        metavar="PATH",
        help="Input Word document (.docx file)",
    )

    parser.add_argument(
        "--input-pptx",
        type=str,
        dest="input_pptx",
        metavar="PATH",
        help="Input PowerPoint file (.pptx file)",
    )
    parser.add_argument(
        "--output-folder",
        type=str,
        dest="output_folder",
        metavar="PATH",
        help="Output folder for converted files",
    )

    # Template files
    parser.add_argument(
        "--template-pptx",
        type=str,
        dest="template_pptx",
        metavar="PATH",
        help="PowerPoint template file",
    )

    parser.add_argument(
        "--template-docx",
        type=str,
        dest="template_docx",
        metavar="PATH",
        help="Word template file",
    )

    # Processing options
    parser.add_argument(
        "--chunk-type",
        type=str,
        dest="chunk_type",
        choices=["paragraph", "page", "heading_flat", "heading_nested"],
        help="How to chunk the document into slides (default: paragraph)",
    )

    parser.add_argument(
        "--direction",
        type=str,
        choices=["docx2pptx", "pptx2docx"],
        help="Conversion direction (default: docx2pptx)",
    )
    # Page range integers
    parser.add_argument(
        "--range-start",
        metavar="N",
        type=int,
        help="EXPERIMENTAL: Specify the page or slide from input to start with. Inclusive. Approximate. For finer control, make a copy of your input file that only includes the excerpt you want processed.",
    )
    parser.add_argument(
        "--range-end",
        metavar="N",
        type=int,
        help="EXPERIMENTAL: Specify the page or slide from input to end on. Inclusive. Approximate. For finer control, make a copy of your input file that only includes the excerpt you want processed.",
    )

    # Boolean flags - experimental formatting
    formatting_group = parser.add_mutually_exclusive_group()
    formatting_group.add_argument(
        "--experimental-formatting",
        action="store_true",
        dest="experimental_formatting_on",
        help="Enable experimental formatting features (default: enabled)",
    )
    formatting_group.add_argument(
        "--no-experimental-formatting",
        action="store_false",
        dest="experimental_formatting_on",
        help="Disable experimental formatting features",
    )

    # Boolean flags - comments
    comments_group = parser.add_mutually_exclusive_group()
    comments_group.add_argument(
        "--display-comments",
        action="store_true",
        dest="display_comments",
        help="Display comments in speaker notes (default: enabled)",
    )
    comments_group.add_argument(
        "--no-display-comments",
        action="store_false",
        dest="display_comments",
        help="Do not display comments",
    )

    # Boolean flags - comments sorting
    comments_sort_group = parser.add_mutually_exclusive_group()
    comments_sort_group.add_argument(
        "--comments-sort-by-date",
        action="store_true",
        dest="comments_sort_by_date",
        help="Sort comments by date (default: enabled)",
    )
    comments_sort_group.add_argument(
        "--no-comments-sort-by-date",
        action="store_false",
        dest="comments_sort_by_date",
        help="Do not sort comments by date",
    )

    # Boolean flags - comments metadata
    comments_metadata_group = parser.add_mutually_exclusive_group()
    comments_metadata_group.add_argument(
        "--comments-keep-author-and-date",
        action="store_true",
        dest="comments_keep_author_and_date",
        help="Keep author and date in comments (default: enabled)",
    )
    comments_metadata_group.add_argument(
        "--no-comments-keep-author-and-date",
        action="store_false",
        dest="comments_keep_author_and_date",
        help="Do not keep author and date in comments",
    )

    # Boolean flags - footnotes
    footnotes_group = parser.add_mutually_exclusive_group()
    footnotes_group.add_argument(
        "--display-footnotes",
        action="store_true",
        dest="display_footnotes",
        help="Display footnotes in speaker notes (default: enabled)",
    )
    footnotes_group.add_argument(
        "--no-display-footnotes",
        action="store_false",
        dest="display_footnotes",
        help="Do not display footnotes",
    )

    # Boolean flags - endnotes
    endnotes_group = parser.add_mutually_exclusive_group()
    endnotes_group.add_argument(
        "--display-endnotes",
        action="store_true",
        dest="display_endnotes",
        help="Display endnotes in speaker notes (default: enabled)",
    )
    endnotes_group.add_argument(
        "--no-display-endnotes",
        action="store_false",
        dest="display_endnotes",
        help="Do not display endnotes",
    )

    # Boolean flags - metadata preservation
    metadata_group = parser.add_mutually_exclusive_group()
    metadata_group.add_argument(
        "--preserve-metadata",
        "--metadata",
        action="store_true",
        dest="preserve_docx_metadata_in_speaker_notes",  # dest is an argparse parameter that tells argparse what attribute name to use when storing the value
        help="Preserve docx metadata in speaker notes (default: enabled)",
    )
    metadata_group.add_argument(
        "--no-preserve-metadata",
        "--no-metadata",
        action="store_false",
        dest="preserve_docx_metadata_in_speaker_notes",
        help="Do not preserve docx metadata",
    )

    # Validate args match config fields
    _validate_args_match_config(parser)

    return parser.parse_args()


def build_config_from_args(args: argparse.Namespace) -> UserConfig:
    """
    Build UserConfig from parsed arguments with proper priority.

    Priority order (highest to lowest):
    1. CLI arguments (if explicitly provided)
    2. Config file values (if --config provided)
    3. UserConfig defaults

    Args:
        args: Parsed command line arguments

    Returns:
        UserConfig instance with all values set
    """
    # Start with demo samples, config file, or defaults
    if args.demo_round_trip:
        log.info(
            "Roundtrip Demo requested; populating input fields with sample defaults."
        )
        cfg = UserConfig().with_defaults()
        cfg.enable_all_options()
        log.debug(
            f"Is Preserve Metadata enabled? {cfg.preserve_docx_metadata_in_speaker_notes}"
        )
        # Early return for demos
        return cfg
    elif args.demo_docx2pptx:
        log.info(
            "docx2pptx demo requested; populating input fields with sample defaults."
        )
        cfg = UserConfig().for_demo(direction=PipelineDirection.DOCX_TO_PPTX)
        return cfg
    elif args.demo_pptx2docx:
        log.info(
            "pptx2docx demo requested; populating input fields with sample defaults."
        )
        cfg = UserConfig().for_demo(direction=PipelineDirection.PPTX_TO_DOCX)
        return cfg
    elif args.config:
        config_path = Path(args.config)
        log.info(f"Loading config from {config_path}")
        cfg = UserConfig.from_toml(config_path)
    else:
        cfg = UserConfig()

    # Override with CLI args (only if explicitly provided)
    # We need to check if the arg was actually provided vs just being the default

    # For string args, check if they're not None
    if args.input_docx is not None:
        cfg.input_docx = args.input_docx
        cfg.direction = PipelineDirection.DOCX_TO_PPTX
    if args.input_pptx is not None:
        cfg.input_pptx = args.input_pptx
        cfg.direction = PipelineDirection.PPTX_TO_DOCX
    if args.output_folder is not None:
        cfg.output_folder = args.output_folder
    if args.template_pptx is not None:
        cfg.template_pptx = args.template_pptx
    if args.template_docx is not None:
        cfg.template_docx = args.template_docx

    if args.range_start is not None:
        cfg.range_start = args.range_start  # argparse already validated it's an int
    if args.range_end is not None:
        cfg.range_end = args.range_end

    # For enums, check if they're not None
    if args.chunk_type is not None:
        cfg.chunk_type = ChunkType(args.chunk_type)
    if args.direction is not None:
        cfg.direction = PipelineDirection(args.direction)

    # For booleans, check if they were explicitly set
    # argparse sets these to None if not provided, or True/False if provided
    if args.experimental_formatting_on is not None:
        cfg.experimental_formatting_on = args.experimental_formatting_on
    if args.display_comments is not None:
        cfg.display_comments = args.display_comments
    if args.comments_sort_by_date is not None:
        cfg.comments_sort_by_date = args.comments_sort_by_date
    if args.comments_keep_author_and_date is not None:
        cfg.comments_keep_author_and_date = args.comments_keep_author_and_date
    if args.display_footnotes is not None:
        cfg.display_footnotes = args.display_footnotes
    if args.display_endnotes is not None:
        cfg.display_endnotes = args.display_endnotes
    if args.preserve_docx_metadata_in_speaker_notes is not None:
        cfg.preserve_docx_metadata_in_speaker_notes = (
            args.preserve_docx_metadata_in_speaker_notes
        )

    # Validate config
    cfg.validate()

    return cfg


# TODO: This should never impact users. We need to gate this with DEBUG_MODE
def _validate_args_match_config(parser: argparse.ArgumentParser) -> None:
    """
    Ensure all UserConfig fields have corresponding CLI arguments.

    This validation catches cases where someone adds a field to UserConfig
    but forgets to add the corresponding CLI argument (or vice versa).

    Raises:
        RuntimeError: If there's a mismatch between config fields and CLI args
    """
    # Get all config field names
    config_fields = {f.name for f in fields(UserConfig)}

    # Get all arg destination names from parser
    # (argparse converts --input-docx to input_docx via arg.dest instead of using arg aliases)
    arg_names = set()
    excluded_args = [
        "help",
        "config",
        "demo_run",
        "demo_round_trip",
        "demo_pptx2docx",
        "demo_docx2pptx",
    ]
    for action in parser._actions:
        # Skip special argparse actions
        if action.dest not in excluded_args:  # exclude --config, --help, etc.
            arg_names.add(action.dest)

    # Find mismatches by getting the difference of the sets
    missing_in_args = config_fields - arg_names
    extra_in_args = arg_names - config_fields

    if missing_in_args:
        log.error(
            "UserConfig fields must have corresponding arg added to cli.parse_args() to ensure parity between interfaces."
        )
        raise RuntimeError(
            f"CLI arguments missing for UserConfig fields: {missing_in_args}\n"
            "These config fields need corresponding arguments added to parse_args()"
        )

    if extra_in_args:
        log.error(
            "We detected unexpected CLI args that do not match UserConfig fields. New argparse fields must either: "
            "1) Have a corresponding configuration option field also added to internals.config.UserConfig() for interface parity, or "
            "2) Be truly CLI interface-specific; e.g., we have --config to pass in a config file to the CLI, and --help to explain the app. "
            "In case 2, you must add the arg to the excluded_args list in _validate_args_match_config()"
        )
        raise RuntimeError(
            f"CLI arguments don't match UserConfig fields: {extra_in_args}\n"
            "Either remove these CLI args or add corresponding fields to UserConfig"
        )


def main() -> None:
    """Development entry point - run CLI directly with `python -m manuscript2slides.cli`"""
    from manuscript2slides import startup

    log = startup.initialize_application()
    try:
        run()
    except Exception:
        log.exception("Fatal error in CLI")
        raise


if __name__ == "__main__":
    main()
