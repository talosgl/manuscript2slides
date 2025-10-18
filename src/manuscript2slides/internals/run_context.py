"""Process-global run ID management.

manuscript2slides uses a single run ID per process to tag logs and output files.
- By default, `get_run_id()` lazily generates an 8-character hex string on first use.
- The run ID never changes for the lifetime of the process.
- Thread-safe: protected by a lock to avoid races.
"""

from __future__ import annotations

import os
import threading
import uuid

# TODO: Separate _run_id into: and _pipeline_run_id and _session_id
# Now that I'm looking at the logs for the UI, I can see that the run_id is going to be the same for any UI session. It's not going to be
# per-pipeline-run, it's going to be per-UI-run. When we were doing CLI, those were the same things, but not with UI. You could, presumably, leave
# the UI open for days, and run it dozens of times, with the same run_id. Dangit!


# Module-level state: one run ID for the entire program lifetime
_run_id: str | None = None  # We start with value as None, to mean "not yet generated"
_id_lock = threading.Lock()


def seed_run_id(value: str) -> None:
    """
    Seed the process-global run ID before it is generated.

    Has no effect if the run ID is already set.
    Useful for testing or sharing IDs across processes.

    Args:
        value: The run ID to set (should be short and unique, e.g., 8 hex chars).
    """
    # Modify the module-level variable; do not create a local one
    global _run_id

    # Lock this section so only one thread can run it at a time; prevents race condition bugs
    with _id_lock:

        # Only set it if it hasn't been set yet
        if _run_id is None:

            # Actually set the value
            _run_id = value


def get_run_id() -> str:
    """
    Return the process-global run ID, generating it if necessary.

    Resolution order:
    1. Already-seeded value (via `seed_run_id()`).
    2. Environment variable `MANUSCRIPT2SLIDES_RUN_ID`.
    3. Fresh random 8-character hex string.

    Returns:
        str: The run ID for this process.
    """
    global _run_id

    # First check (outside lock): Fast path. If ID already exists, just return it. No lock needed = faster.
    if _run_id is None:
        with _id_lock:
            # Second check (inside lock): Safety. Multiple threads might pass the first check simultaneously. The lock ensures only one generates the ID.
            if _run_id is None:
                # Check environment variable first; if that's empty, perform the right-side of the `or` and generate random UUID
                _run_id = (
                    os.environ.get("MANUSCRIPT2SLIDES_RUN_ID")
                    or uuid.uuid4().hex[:8]  # Reduce to 8 char for brevity/readability
                )
    return _run_id
