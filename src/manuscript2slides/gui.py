"""GUI Interface logic."""

from __future__ import annotations

from manuscript2slides.internals.config.define_config import (
    UserConfig,
    ChunkType,
)

from manuscript2slides import startup
from manuscript2slides.orchestrator import (
    run_pipeline,
)  # we'll need this later to replace run_roundtrip_test

import logging

log = logging.getLogger("manuscript2slides")


def run() -> None:
    """Run GUI interface. Assumes startup.initialize_application() already called."""

    # Create config with defaults
    cfg = UserConfig()

    # TODO: Use a GUI to populate the config with user values

    # Validate config shape
    cfg.validate()

    # === Pipeline testing
    # TODO: Replace with simple run_pipeline(cfg) once UI is ready
    cfg.chunk_type = ChunkType.HEADING_FLAT

    # Temporary: Run round-trip test for development/testing
    from manuscript2slides.orchestrator import run_roundtrip_test

    run_roundtrip_test(cfg)
