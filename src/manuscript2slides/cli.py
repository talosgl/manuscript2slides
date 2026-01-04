"""CLI Interface Logic (argparse etc)"""

# region imports
import argparse
import logging
import sys
from dataclasses import fields
from pathlib import Path

from manuscript2slides import __version__
from manuscript2slides.internals.define_config import (
    ChunkType,
    PipelineDirection,
    UserConfig,
)
from manuscript2slides.orchestrator import run_pipeline, run_roundtrip_test
from manuscript2slides.utils import get_debug_mode

# endregion

log = logging.getLogger("manuscript2slides")


# region run
def run() -> None:
    """Run CLI interface. Assumes startup.initialize_application() was already called.

    Called by:
        # After pip install
        manuscript2slides-cli

        # From source code
        python -m manuscript2slides.cli
    """

    # (Define and) parse user-passed-in command line arguments for this app
    args = parse_args()

    # Build config from args (with proper prioritization CLI args > config file > defaults)
    cfg = build_config_from_args(args)

    # If round-trip demo was enabled, run that special orchestrator
    if args.demo_round_trip:
        log.info("Running round-trip test with sample files.")
        run_roundtrip_test(cfg)
    else:
        run_pipeline(cfg)


# endregion


# region parse_args()
def parse_args() -> argparse.Namespace:
    """
    Parse command line arguments.

    Returns argparse.Namespace with all the UserConfig fields as attributes.
    Validates that all config fields hav corresponding CLI arguments.
    """

    # Create the argparse object in memory
    parser = argparse.ArgumentParser(
        prog="manuscript2slides-cli",
        description="Convert text content from Word docx to PowerPoint pptx slides and vice versa",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
------------------------------------------------------------------

Examples:
  # Use the GUI
  manuscript2slides

  # Use the CLI:
  manuscript2slides-cli --help
  manuscript2slides-cli --version

  # See a demo run with sample files
  manuscript2slides-cli --demo-docx2pptx
  manuscript2slides-cli --demo-pptx2docx
  manuscript2slides-cli --demo-round-trip

  # Quick conversion of a real file with default options
  manuscript2slides-cli --input-docx path/to/your_manuscript.docx

  # Use config file
  manuscript2slides-cli --config path/to/my_settings.toml

  # Override config file settings
  manuscript2slides-cli --config settings.toml --input-pptx path/to/your_slides.pptx

Environment Variables:
  MANUSCRIPT2SLIDES_DEBUG=true     Enable debug mode (extra validation/logging)
  MANUSCRIPT2SLIDES_USER_DIR=path  Override default user directory
------------------------------------------------------------------
""",
    )

    # Add version flag
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    # Add all our arguments to the argparse object we just made

    # Args for running dry runs
    parser.add_argument(
        "--demo-docx2pptx",
        "--demo_docx2pptx",
        action="store_true",
        dest="demo_docx2pptx",
        help=(
            "Run a demonstration conversion from a sample Word document to PowerPoint slides. "
            "Ignores other CLI options and uses built-in example files."
        ),
    )
    parser.add_argument(
        "--demo-pptx2docx",
        "--demo_pptx2docx",
        action="store_true",
        dest="demo_pptx2docx",
        help="Run a demonstration conversion from sample PowerPoint slides to a Word document. "
        "Ignores other CLI options and uses built-in example files.",
    )

    # Arg for running dry run
    parser.add_argument(
        "--demo-round-trip",
        "--demo_round_trip",
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
        "--input_docx",
        type=str,
        dest="input_docx",
        metavar="PATH",
        help="Input Word document (.docx file)",
    )

    parser.add_argument(
        "--input-pptx",
        "--input_pptx",
        type=str,
        dest="input_pptx",
        metavar="PATH",
        help="Input PowerPoint file (.pptx file)",
    )
    parser.add_argument(
        "--output-folder",
        "--output_folder",
        "--output",
        "--output-dir",
        "-o",
        type=str,
        dest="output_folder",
        metavar="PATH",
        help="Output folder for converted files",
    )

    # Template files
    parser.add_argument(
        "--template-pptx",
        "--template_pptx",
        type=str,
        dest="template_pptx",
        metavar="PATH",
        help="PowerPoint template file",
    )

    parser.add_argument(
        "--template-docx",
        "--template_docx",
        type=str,
        dest="template_docx",
        metavar="PATH",
        help="Word template file",
    )

    # Processing options
    parser.add_argument(
        "--chunk-type",
        "--chunk_type",
        "--chunk",
        "--chunk-by",
        "--chunk_by",
        "-c",
        type=str,
        dest="chunk_type",
        choices=[
            "paragraph",
            "page",
            "heading_flat",
            "heading_nested",
            "heading",  # heading is an alias for heading_flat
        ],
        help="How to chunk the document into slides (default: paragraph)",
    )

    # Page range integers
    parser.add_argument(
        "--range-start",
        "--range_start",
        "--start",
        dest="range_start",
        metavar="N",
        type=int,
        help="EXPERIMENTAL: Specify the page or slide from input to start with. Inclusive. Approximate. For finer control, make a copy of your input file that only includes the excerpt you want processed.",
    )
    parser.add_argument(
        "--range-end",
        "--range_end",
        "--end",
        "--range-stop",
        "--range_stop",
        "--stop",
        dest="range_end",
        metavar="N",
        type=int,
        help="EXPERIMENTAL: Specify the page or slide from input to end on. Inclusive. Approximate. For finer control, make a copy of your input file that only includes the excerpt you want processed.",
    )

    # Boolean flags - experimental formatting
    formatting_group = parser.add_mutually_exclusive_group()
    formatting_group.add_argument(
        "--experimental-formatting",
        action="store_const",
        const=True,
        dest="experimental_formatting_on",
        help="Enable experimental formatting features (default: enabled)",
    )
    formatting_group.add_argument(
        "--no-experimental-formatting",
        action="store_const",
        const=False,
        dest="experimental_formatting_on",
        help="Disable experimental formatting features",
    )

    # Boolean flags - comments
    comments_group = parser.add_mutually_exclusive_group()
    comments_group.add_argument(
        "--display-comments",
        action="store_const",
        const=True,
        dest="display_comments",
        help="Display comments in speaker notes (default: disabled)",
    )
    comments_group.add_argument(
        "--no-display-comments",
        action="store_const",
        const=False,
        dest="display_comments",
        help="Do not display comments",
    )

    # Boolean flags - comments sorting
    comments_sort_group = parser.add_mutually_exclusive_group()
    comments_sort_group.add_argument(
        "--comments-sort-by-date",
        action="store_const",
        const=True,
        dest="comments_sort_by_date",
        help="Sort comments by date (default: enabled)",
    )
    comments_sort_group.add_argument(
        "--no-comments-sort-by-date",
        action="store_const",
        const=False,
        dest="comments_sort_by_date",
        help="Do not sort comments by date",
    )

    # Boolean flags - comments metadata
    comments_metadata_group = parser.add_mutually_exclusive_group()
    comments_metadata_group.add_argument(
        "--comments-keep-author-and-date",
        action="store_const",
        const=True,
        dest="comments_keep_author_and_date",
        help="Keep author and date in comments (default: enabled)",
    )
    comments_metadata_group.add_argument(
        "--no-comments-keep-author-and-date",
        action="store_const",
        const=False,
        dest="comments_keep_author_and_date",
        help="Do not keep author and date in comments",
    )

    # Boolean flags - footnotes
    footnotes_group = parser.add_mutually_exclusive_group()
    footnotes_group.add_argument(
        "--display-footnotes",
        action="store_const",
        const=True,
        dest="display_footnotes",
        help="Display footnotes in speaker notes (default: disabled)",
    )
    footnotes_group.add_argument(
        "--no-display-footnotes",
        action="store_const",
        const=False,
        dest="display_footnotes",
        help="Do not display footnotes",
    )

    # Boolean flags - endnotes
    endnotes_group = parser.add_mutually_exclusive_group()
    endnotes_group.add_argument(
        "--display-endnotes",
        action="store_const",
        const=True,
        dest="display_endnotes",
        help="Display endnotes in speaker notes (default: disabled)",
    )
    endnotes_group.add_argument(
        "--no-display-endnotes",
        action="store_const",
        const=False,
        dest="display_endnotes",
        help="Do not display endnotes",
    )

    # Boolean flags - metadata preservation
    metadata_group = parser.add_mutually_exclusive_group()
    metadata_group.add_argument(
        "--preserve-metadata",
        "--metadata",
        action="store_const",
        const=True,
        dest="preserve_docx_metadata_in_speaker_notes",  # dest is an argparse parameter that tells argparse what attribute name to use when storing the value
        help="Preserve docx metadata in speaker notes (default: disabled)",
    )
    metadata_group.add_argument(
        "--no-preserve-metadata",
        "--no-metadata",
        action="store_const",
        const=False,
        dest="preserve_docx_metadata_in_speaker_notes",
        help="Do not preserve docx metadata",
    )

    if get_debug_mode():
        # Validate args match config fields
        _validate_args_match_config(parser)

    args = parser.parse_args()

    # If no args were passed in, act as if --help was passed and show help.
    # len(sys.argv) == 1 will be true if only the script name was passed (`python -m manuscript2slides.cli` in dev, or `manuscript2slides-cli` via pip install)
    if len(sys.argv) == 1:
        # We use stderr so the message is visible even if stdout is piped to a file
        print("Error: No arguments provided. Showing help.\n", file=sys.stderr)
        log.warning(
            "No args passed to CLI. Showing help. (Try passing `--demo-round-trip` or `--demo-docx2pptx` to see the pipeline in a dry run.)"
        )
        parser.print_help(sys.stderr)
        sys.exit(2)  # Standard argparse exit code for usage errors

    return args


# endregion


# region build_config_from_args
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
    #"""
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
        cfg = UserConfig().for_demo(requested_direction=PipelineDirection.DOCX_TO_PPTX)
        return cfg
    elif args.demo_pptx2docx:
        log.info(
            "pptx2docx demo requested; populating input fields with sample defaults."
        )
        cfg = UserConfig().for_demo(requested_direction=PipelineDirection.PPTX_TO_DOCX)
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
        cfg.input_pptx = None  # Clear the opposite input

    if args.input_pptx is not None:
        cfg.input_pptx = args.input_pptx
        cfg.input_docx = None  # Clear the opposite input

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
        cfg.chunk_type = ChunkType.from_string(args.chunk_type)

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

    # Recreate config to ensure __post_init__ runs with final values
    final_cfg = UserConfig(**{f.name: getattr(cfg, f.name) for f in fields(cfg)})

    # Validate config
    final_cfg.validate()

    return final_cfg


# endregion


# region _validate_args_match_config
def _validate_args_match_config(parser: argparse.ArgumentParser) -> None:
    """
    Ensure all UserConfig fields have corresponding CLI arguments.

    This validation catches cases where someone adds a field to UserConfig
    but forgets to add the corresponding CLI argument (or vice versa).

    Raises:
        RuntimeError: If there's a mismatch between config fields and CLI args
    """

    if not get_debug_mode():
        # Return early if debug_mode is turned off.
        # This function should never crash the app for users.
        return

    # Get all config field names
    config_fields = {f.name for f in fields(UserConfig)}

    # Get all arg destination names from parser
    # (argparse converts --input-docx to input_docx via arg.dest instead of using arg aliases)
    arg_names = set()
    excluded_args = {
        "help",
        "version",
        "config",
        "demo_run",
        "demo_round_trip",
        "demo_pptx2docx",
        "demo_docx2pptx",
    }

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


# endregion


# region main
def main() -> None:
    """CLI entry point.

    After pip install:
        manuscript2slides-cli

    From source (dev):
        python -m manuscript2slides.cli
    """
    from manuscript2slides import startup

    log = startup.initialize_application()
    try:
        run()
    except Exception:
        log.exception("Fatal error in CLI")
        raise


if __name__ == "__main__":
    main()
# endregion
