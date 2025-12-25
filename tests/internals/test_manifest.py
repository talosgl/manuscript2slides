"""Test the manifest system.

The manifest should never crash the app - it's a "nice to have" feature that records
metadata about pipeline runs. Tests focus on ensuring it writes correct data and fails
gracefully.
"""

import pytest
from pathlib import Path
from manuscript2slides.internals.manifest import RunManifest, MANIFEST_VERSION
from manuscript2slides.internals.define_config import UserConfig, PipelineDirection
import json


def test_manifest_creates_file_on_start_and_has_required_fields(
    sample_d2p_cfg: UserConfig,
    temp_output_dir: Path,
) -> None:
    """Verify manifest creates a JSON file when start() is called."""

    # Create manifest in temp directory
    run_id = "test_run_creates_with_fields"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)

    # Before start(), file shouldn't exist
    assert not manifest.manifest_path.exists()

    # After start(), file should exist
    manifest.manifest_path = temp_output_dir / f"run_{run_id}_manifest.json"
    manifest.start()

    assert manifest.manifest_path.exists()
    assert manifest.manifest_path.is_file()


    # Read the manifest file
    with open(manifest.manifest_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Check required fields
    assert data["manifest_version"] == MANIFEST_VERSION
    assert data["run_id"] == run_id
    assert data["status"] == "running"
    assert data["direction"] == PipelineDirection.DOCX_TO_PPTX.value
    assert "start_time" in data
    assert "environment" in data
    assert "config" in data
    assert data["end_time"] is None
    assert data["duration_seconds"] is None


def test_manifest_updates_on_completion(
    sample_d2p_cfg: UserConfig,
    temp_output_dir: Path,
) -> None:
    """Verify manifest updates correctly when complete() is called."""
    run_id = "test_run_update_on_complete"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)
    manifest.manifest_path = temp_output_dir / f"run_{run_id}_manifest.json"
    manifest.start()

    # Mark as complete
    output_path = temp_output_dir / "output.pptx"
    manifest.complete(output_path)

    # Read updated manifest
    with open(manifest.manifest_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Check completion fields
    assert data["status"] == "success"
    assert data["output_path"] == str(output_path)
    assert data["end_time"] is not None
    assert data["duration_seconds"] is not None
    assert isinstance(data["duration_seconds"], (int, float))
    assert data["duration_seconds"] >= 0


def test_manifest_updates_on_failure(
    sample_d2p_cfg: UserConfig,
    temp_output_dir: Path,
) -> None:
    """Verify manifest updates correctly when fail() is called."""
    run_id = "test_run_update_on_fail"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)
    manifest.manifest_path = temp_output_dir / f"run_{run_id}_manifest.json"
    manifest.start()

    # Mark as failed
    test_error = ValueError("Test error message")
    manifest.fail(test_error)

    # Read updated manifest
    with open(manifest.manifest_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Check failure fields
    assert data["status"] == "fail"
    assert data["error"] == "Test error message"
    assert data["error_type"] == "ValueError"
    assert data["end_time"] is not None
    assert data["duration_seconds"] is not None
    assert isinstance(data["duration_seconds"], (int, float))


def test_manifest_environment_info(
    sample_d2p_cfg: UserConfig,
) -> None:
    """Verify environment information is captured."""
    run_id = "test_run_env_info"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)

    env = manifest.manifest["environment"]

    # Check environment has expected fields
    assert "python_version" in env
    assert "platform" in env
    assert "platform_release" in env
    assert "app_version" in env

    # Basic sanity checks
    assert len(env["python_version"]) > 0
    assert isinstance(env["platform"], str) and len(env["platform"]) > 0


@pytest.mark.parametrize("cfg_fixture,expected_name,run_id", [
    ("sample_d2p_cfg", "run_docx2pptx_pipeline", "test_d2p_name" ),
    ("sample_p2d_cfg", "run_pptx2docx_pipeline", "test_p2d_name")
])
def test_manifest_gets_correct_pipeline_name(
    cfg_fixture: str, expected_name: str, run_id: str, request: pytest.FixtureRequest, 
) -> None:
    """Verify pipeline name is correct for docx2pptx direction."""   
    cfg = request.getfixturevalue(cfg_fixture)
 
    manifest = RunManifest(cfg, run_id=run_id)

    assert manifest._get_pipeline_name() == expected_name
    assert manifest.manifest["pipeline_name"] == expected_name


def test_manifest_handles_missing_directory_gracefully_and_logs(
    sample_d2p_cfg: UserConfig,
    temp_output_dir: Path,
    caplog: pytest.LogCaptureFixture
) -> None:
    """Verify manifest doesn't crash if output directory doesn't exist.

    This is important because the manifest is a 'nice to have' feature and
    should never crash the main application.
    """
    run_id = "test_run_missing_dir"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)

    # Point to a non-existent directory
    non_existent_dir = temp_output_dir / "does_not_exist" / "subdirectory"
    manifest.manifest_path = non_existent_dir / f"run_{run_id}_manifest.json"

    # This should not raise an exception - it should fail silently
    # (the actual implementation logs an error but doesn't crash)
    manifest.start()

    # File won't exist, but the call shouldn't have crashed
    assert not manifest.manifest_path.exists()

    # Verify error was logged (not crashed)
    assert "Failed to write manifest" in caplog.text


def test_manifest_duration_calculation(
    sample_d2p_cfg: UserConfig,
    temp_output_dir: Path,
) -> None:
    """Verify duration is calculated correctly."""
    import time

    run_id = "test_run_duration_calc"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)
    manifest.manifest_path = temp_output_dir / f"run_{run_id}_manifest.json"
    manifest.start()

    # Wait a tiny bit
    time.sleep(0.1)

    # Complete the run
    output_path = temp_output_dir / "output.pptx"
    manifest.complete(output_path)

    # Duration should be at least 0.1 seconds
    assert manifest.duration is not None
    assert manifest.duration > 0 
    assert manifest.end_time is not None
    assert manifest.end_time > manifest.start_time


def test_manifest_json_format_and_encoding(
    sample_d2p_cfg: UserConfig,
    temp_output_dir: Path,
) -> None:
    """Verify the manifest file is valid, well-formatted JSON with proper UTF-8 encoding."""
    run_id = "test_run_json_format"
    manifest = RunManifest(sample_d2p_cfg, run_id=run_id)
    manifest.manifest_path = temp_output_dir / f"run_{run_id}_manifest.json"
    manifest.start()

    # Read raw bytes (rb)
    with open(manifest.manifest_path, 'rb') as f:
        raw_bytes = f.read()

    # Case: Verify it's valid UTF-8 (and assign to a variable for later)
    content = raw_bytes.decode('utf-8')  # Will raise if not valid UTF-8

    # Case: Should be able to parse as JSON without errors
    # (This is covered in other tests, but replicated here in case
    # it's useful later for debugging to have it narrowed.)
    data = json.loads(content)  # Will raise if invalid JSON
    assert isinstance(data, dict)

    # Case: Check if the JSON is formatted with proper indentation (pretty printed instead of minified)
    # Check if it contains 2-space indents
    assert '  ' in content
    # Verify it's not minified to 1 line
    assert '\n' in content

    # Reserialize with indent=2 and compare
    expected = json.dumps(data, indent=2, ensure_ascii=False)
    actual = content.strip()

    assert actual == expected