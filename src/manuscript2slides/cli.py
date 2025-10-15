"""CLI Interface Logic (argparse etc)"""

from __future__ import annotations

from manuscript2slides.utils import setup_console_encoding
from manuscript2slides.internals.config.define_config import (
    UserConfig,
    ChunkType,
)
from manuscript2slides.orchestrator import (
    run_pipeline,
)  # we'll need this later to replace run_roundtrip_test

import logging

log = logging.getLogger("manuscript2slides")


def run() -> None:
    """Run CLI interface. Assumes startup.initialize_application() was already called."""

    # ==== Logic that will be in both CLI & GUI in some form

    # Create config with defaults
    cfg = UserConfig()

    # == TODOs for later, when I care to expand CLI interface:
    # TODO add capability to populate/override defaults with user-provided CLI Args
    # TODO add capability to populate from toml config file (but prioritize CLI args)
    # TODO add some 3-tier merging system that prioritizes: CLI args > toml config file > UserConfig class defaults
    # ==

    cfg.validate()

    # == Pipeline testing.
    # TODO: Replace with simple run_pipeline(cfg) once CLI is ready
    cfg.chunk_type = ChunkType.HEADING_FLAT

    # Temporary: Run round-trip test for development/testing
    from manuscript2slides.orchestrator import run_roundtrip_test

    run_roundtrip_test(cfg)
