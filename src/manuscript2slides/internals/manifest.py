"""Track and record metadata for pipeline runs."""

from __future__ import annotations

import json
import logging

import platform
import sys
from datetime import datetime
from pathlib import Path

from manuscript2slides.internals.config.define_config import (
    UserConfig,
    PipelineDirection,
)
from manuscript2slides.internals.run_context import get_session_id
from manuscript2slides.internals.paths import user_manifests_dir, user_log_dir_path

log = logging.getLogger("manuscript2slides")

# Schema version for this implementation
MANIFEST_VERSION = "1.0"


class RunManifest:
    """Tracks and records metadata for a pipeline run."""

    def __init__(self, cfg: UserConfig, run_id: str) -> None:
        """Creates a run manifest object in memory with initial fields. Caller must immediately call .start()."""
        self.cfg = cfg
        self.run_id = run_id
        self.start_time: datetime = datetime.now()
        self.manifest_path = user_manifests_dir() / f"run_{self.run_id}_manifest.json"
        self.manifest = self._build_manifest()

        # Initialize end_time and duration attributes (they are None until completed/failed)
        # We do that for these, but not for "status", because we'll be using these for runtime operations later,
        # while status is simply used for writing to disk.
        self.end_time: datetime | None = None
        self.duration: float | None = None

    def start(self) -> None:
        """Write initial manifest to disk"""
        # By common  Python convention, __init__ shouldn't do I/O, so we separate this step from the constructor.

        # Update any fields we need to since constructor was called
        self.manifest["status"] = (
            "running"  # Not needed if we set in constructor, but I haven't decided & overwriting seems harmless
        )

        log.info(
            f"Writing initial manifest to disk with status = running, at {self.manifest_path}"
        )
        self._write_manifest()

    def complete(self, output_path: Path) -> None:
        """Update manifest on success"""
        self._get_time_stats()

        # Update the object in memory
        self.manifest["status"] = "success"
        self.manifest["end_time"] = self.end_time.isoformat() if self.end_time else None
        self.manifest["duration_seconds"] = self.duration
        self.manifest["output_path"] = str(output_path)

        self._write_manifest()
        log.info(f"Updated manifest: success, at {self.manifest_path}")

    def fail(self, error: Exception) -> None:
        """Update manifest on failure with error information."""
        self._get_time_stats()

        # Update dict in memory with status="fail", error=...
        self.manifest["status"] = "fail"
        self.manifest["error"] = str(error)
        self.manifest["error_type"] = type(error).__name__
        self.manifest["end_time"] = self.end_time.isoformat() if self.end_time else None
        self.manifest["duration_seconds"] = self.duration

        # Write to disk
        self._write_manifest()
        log.error(f"Updated manifest ({self.manifest_path}): failed - {error}")

    # Private methods
    def _build_manifest(self) -> dict:
        """Build manifest structure. Separating into its own method in case we want to extend in a v2 later."""

        manifest = {
            "manifest_version": MANIFEST_VERSION,
            "run_id": self.run_id,
            "session_id": get_session_id(),
            "environment": self._get_environment_info(),
            "start_time": self.start_time.isoformat(),
            "end_time": None,
            "duration_seconds": None,
            "direction": self.cfg.direction.value,
            "pipeline_name": self._get_pipeline_name(),
            "input_file": str(self.cfg.get_input_file()),
            "output_folder": str(self.cfg.get_output_folder()),
            "log_path": str(user_log_dir_path()),
            "config": self.cfg.config_to_dict(),
            "error": None,
            "error_type": None,
        }

        return manifest

    # Helpers
    def _write_manifest(self) -> None:
        """Write manifest to disk."""
        if not self.manifest_path:
            return

        try:
            with open(self.manifest_path, "w", encoding="utf-8") as f:
                json.dump(self.manifest, f, indent=2)
        except OSError as e:
            log.error(f"Failed to write manifest to {self.manifest_path}: {e}")

    def _get_time_stats(self) -> None:
        self.end_time = datetime.now()
        self.duration = (self.end_time - self.start_time).total_seconds()

    def _get_pipeline_name(self) -> str:
        if self.cfg.direction == PipelineDirection.DOCX_TO_PPTX:
            return "run_docx2pptx_pipeline"
        elif self.cfg.direction == PipelineDirection.PPTX_TO_DOCX:
            return "run_pptx2docx_pipeline"
        else:
            return "unknown_pipeline"

    def _get_environment_info(self) -> dict:
        """Get execution environment information."""
        return {
            "python_version": sys.version.split()[0],
            "platform": platform.system(),
            "platform_release": platform.release(),
            "app_version": self._get_app_version(),
        }

    def _get_app_version(self) -> str:
        """Get manuscript2slides version."""
        try:
            from manuscript2slides import __version__

            return __version__
        except (ImportError, AttributeError):
            return "unknown"
