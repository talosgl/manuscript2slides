"""Process-global execution context management.

Manages two levels of tracking IDs:
- session_id: Generated once per app startup (entire GUI session or CLI invocation)
- pipeline_run_id: Generated fresh for each pipeline execution
"""

from __future__ import annotations

import os
import threading
import uuid

# Module-level state: one session ID for the entire program lifetime
# We start with value as None, to mean "not yet generated"
_session_id: str | None = None

# Unlike, session, pipeline ID will be generated anew on any pipeline-run, which could occur multiple times in a GUI session.
_pipeline_run_id: str | None = None

_session_lock = threading.Lock()
_pipeline_lock = threading.Lock()


# region seed_session_id
def seed_session_id(value: str) -> None:
    """
    Seed the process-global session ID before it is generated.

    Has no effect if the session ID is already set.
    Useful for testing or controlled initialization.

    Args:
        value: The session ID to set (should be short and unique, e.g., 8 hex chars).
    """
    # Modify the module-level variable; do not create a local one
    global _session_id

    # Lock this section so only one thread can run it at a time; prevents race condition bugs
    with _session_lock:
        # Only set it if it hasn't been set yet
        if _session_id is None:
            # Actually set the value
            _session_id = value


# endregion


# region get_session_id
def get_session_id() -> str:
    """
    Return the process-global session ID, generating it if necessary.

    Resolution order:
    1. Already-seeded value (via `seed_session_id()`).
    2. Environment variable `MANUSCRIPT2SLIDES_SESSION_ID`.
        - Allows tests to set a predictable session ID for assertions
        and log correlation without modifying code or calling
        seed_session_id()
        - Allows external systems (CI/CD pipelines, orchestration tools,
        containerized environments) to set a known session ID rather
        than having a random one generated.

    3. Fresh random 8-character hex string.

    Returns:
        str: The session ID for this process.
    """
    global _session_id

    # First check (outside lock): Fast path. If ID already exists, just return it. No lock needed = faster.
    if _session_id is None:
        with _session_lock:
            # Second check (inside lock): Safety. Multiple threads might pass the first check simultaneously. The lock ensures only one generates the ID.
            if _session_id is None:
                # Check environment variable first; if that's empty, perform the right-side of the `or` and generate random UUID
                _session_id = (
                    os.environ.get("MANUSCRIPT2SLIDES_SESSION_ID")
                    or uuid.uuid4().hex[:8]
                )
    return _session_id


# endregion


# region start_pipeline_run
def start_pipeline_run() -> str:
    """
    Generate and set a fresh pipeline run ID.

    Call at the start of each pipeline execution.
    Always generate a new ID, even if one already exists.

    Returns:
        str: The newly generated pipeline run ID.
    """
    global _pipeline_run_id

    with _pipeline_lock:
        # The major difference between this and session ID is: here, we are not checking if this id is None.
        # We're always going to overwrite it when this gets called (assuming threading lets us)
        _pipeline_run_id = uuid.uuid4().hex[:8]

    return _pipeline_run_id


# endregion


# region get_pipeline_run_id
def get_pipeline_run_id() -> str:
    """
    Return the current pipeline run ID.

    Raises:
        RuntimeError: If no pipeline run is active (start_pipeline_run() was not called)

    Returns:
        str: the pipeline run ID for the current execution.
    """
    if _pipeline_run_id is None:
        import logging

        log = logging.getLogger("manuscript2slides")
        log.debug(
            "There is no _pipeline_run_id set yet; returning Unknown. Call start_pipeline_run() at the beginning of pipeline execution."
        )
        return "Unknown"
    return _pipeline_run_id


# endregion


# region seed_pipeline_run_id
def seed_pipeline_run_id(value: str) -> None:
    """
    Seed the pipeline run ID (primarily for testing).

    Unlike normal usage, this allows setting the ID without going through start_pipeline_run().

    Args:
        value: The pipeline run ID to set.
    """
    global _pipeline_run_id
    with _pipeline_lock:
        _pipeline_run_id = value


# endregion
