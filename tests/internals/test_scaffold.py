"""Tests to ensure the manuscript2slides directory structure gets created properly under the users' Documents."""

from pathlib import Path

import pytest

from manuscript2slides.internals import scaffold


# region test ensure_user_scaffold
def test_ensure_user_scaffold_happy_path(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test that scaffold creates expected directory structure."""

    # Mock user_base_dir() to return tmp_path, for both the returning module & calling module.
    # If we don't do it in the calling module (scaffold), then when we store the return in a var,
    # (e.g., `line 48: base = user_base_dir()`), the real return value will be stored and used,
    # (e.g., by line 60) not the mock.
    monkeypatch.setattr(
        "manuscript2slides.internals.paths.user_base_dir", lambda: tmp_path
    )
    monkeypatch.setattr(
        "manuscript2slides.internals.scaffold.user_base_dir", lambda: tmp_path
    )

    # Now when ensure_user_scaffold() runs, all the path functions
    # will create subdirs under tmp_path automatically
    scaffold.ensure_user_scaffold()

    # Assert structure was created
    assert (tmp_path / "README.md").exists()
    assert (tmp_path / "README.md").is_file()  # implied .exists() check
    assert (tmp_path / "input").is_dir()
    assert (tmp_path / "output").is_dir()
    assert (tmp_path / "logs").is_dir()
    assert (tmp_path / "templates").is_dir()
    assert (tmp_path / "configs").is_dir()
    assert (tmp_path / "manifests").is_dir()
    assert (tmp_path / "templates" / "pptx_template.pptx").is_file()
    assert (tmp_path / "templates" / "docx_template.docx").is_file()
    assert (tmp_path / "input" / "sample_doc.docx").is_file()
    assert (tmp_path / "configs" / "sample_config.toml").is_file()


def test_ensure_user_scaffold_does_not_overwrite(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test that calling ensure_user_scaffold() multiple times doesn't overwrite existing files."""

    # Mock user_base_dir in both places
    monkeypatch.setattr(
        "manuscript2slides.internals.paths.user_base_dir", lambda: tmp_path
    )
    monkeypatch.setattr(
        "manuscript2slides.internals.scaffold.user_base_dir", lambda: tmp_path
    )

    # Create scaffold first time
    scaffold.ensure_user_scaffold()

    # Modify the README
    readme_path = tmp_path / "README.md"
    original_content = readme_path.read_text(encoding="utf-8")
    modified_content = "# CUSTOM USER CONTENT - DO NOT OVERWRITE"
    readme_path.write_text(modified_content, encoding="utf-8")

    # Modify a template file
    template_path = tmp_path / "templates" / "pptx_template.pptx"
    original_size = template_path.stat().st_size

    # Write some junk to make it different (not a valid pptx, but that's okay for this test)
    template_path.write_bytes(b"CUSTOM TEMPLATE DATA")
    modified_size = template_path.stat().st_size

    # Call ensure_user_scaffold() again
    scaffold.ensure_user_scaffold()

    # Assert files were NOT overwritten
    assert readme_path.read_text() == modified_content
    assert readme_path.read_text() != original_content

    assert template_path.stat().st_size == modified_size
    assert template_path.stat().st_size != original_size


def test_ensure_scaffold_creates_fallback_readme_when_source_missing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """Test that a fallback README is created when the resource source is missing."""

    # Mock paths
    monkeypatch.setattr(
        "manuscript2slides.internals.paths.user_base_dir", lambda: tmp_path
    )
    monkeypatch.setattr(
        "manuscript2slides.internals.scaffold.user_base_dir", lambda: tmp_path
    )

    # Mock _get_resource_path to return non-existent file
    monkeypatch.setattr(
        "manuscript2slides.internals.scaffold._get_resource_path",
        lambda filename: tmp_path / "nonexistent" / filename,
    )

    # Call ensure_user_scaffold
    scaffold.ensure_user_scaffold()

    # Assert fallback README was created
    readme = tmp_path / "README.md"
    assert readme.is_file()
    content = readme.read_text(encoding="utf-8")
    assert "User folder created automatically" in content  # Fallback text from line 114

    # Assert error was logged
    assert "README template not found" in caplog.text


def test_sample_config_has_valid_toml_and_correct_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test that sample_config.toml has valid TOML and correct paths"""
    # Mock paths
    monkeypatch.setattr(
        "manuscript2slides.internals.paths.user_base_dir", lambda: tmp_path
    )
    monkeypatch.setattr(
        "manuscript2slides.internals.scaffold.user_base_dir", lambda: tmp_path
    )

    # Call ensure_user_scaffold
    scaffold.ensure_user_scaffold()

    # The sample config is at tmp_path / "configs" / "sample_config.toml"
    sample_config_path = tmp_path / "configs" / "sample_config.toml"

    # Make sure the sample config exists and is a file.
    assert sample_config_path.is_file(), (
        f"We couldn't find the sample_config at {sample_config_path}, under {tmp_path}"
    )

    # Read and validate it's valid TOML
    try:
        import tomllib  # Python 3.11+
    except ModuleNotFoundError:
        import tomli as tomllib  # type: ignore[no-redef]  # Python 3.10

    with open(sample_config_path, "rb") as f:
        config_data = tomllib.load(f)

    # Assert paths are based on tmp_path (user-specific absolute paths)
    assert "input_docx" in config_data
    assert config_data["input_docx"].startswith(
        tmp_path.as_posix()
    )  # Path should contain tmp_path
    assert config_data["output_folder"] == (tmp_path / "output").as_posix()
    assert (
        config_data["template_pptx"]
        == (tmp_path / "templates" / "pptx_template.pptx").as_posix()
    )


# endregion

# region test helpers
# Happy path tests for helpers to help narrow down issues when ensure_user_scaffold tests fail


def test_create_readme_happy_path(tmp_path: Path) -> None:
    """Verify we can create some form of readme."""
    readme_path = scaffold._create_readme_if_missing(tmp_path)
    assert readme_path.is_file()


def test_copy_templates_if_missing_happy_path(tmp_path: Path) -> None:
    """Verify we create templates."""
    template_paths_list = scaffold._copy_templates_if_missing(tmp_path)
    for path in template_paths_list:
        assert path.is_file()


def test_copy_samples_if_missing_happy_path(tmp_path: Path) -> None:
    """Verify we create sample docx and pptx."""
    sample_paths_list = scaffold._copy_samples_if_missing(tmp_path)
    for path in sample_paths_list:
        assert path.is_file()


def test_copy_sample_config_if_missing_happy_path(tmp_path: Path) -> None:
    """Verify we create sample config."""
    sample_config_path = scaffold._copy_sample_config_if_missing(tmp_path)
    assert sample_config_path.is_file()


# endregion
