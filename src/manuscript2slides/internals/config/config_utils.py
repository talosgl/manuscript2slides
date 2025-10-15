# internals/config/config_utils.py
"""YAML configuration file parsing utilities."""

import yaml
from pathlib import Path
from manuscript2slides.internals.config.define_config import (
    ChunkType,
    PipelineDirection,
)


# TODO: If/when we get back to supporting on-disk config files, don't use yaml. Refactor to use toml, instead.
def load_yaml_config(path: Path) -> dict:
    """Safe load and parse a YAML config file."""

    # Add this check
    if path.suffix.lower() not in {".yaml", ".yml"}:
        raise ValueError(f"Config file must be .yaml or .yml, got {path.suffix}")

    try:
        with open(path, "r", encoding="utf-8") as f:
            # safe_load prevents arbitrary code execution. Always use this, never yaml.load()
            data = yaml.safe_load(f)
    except yaml.YAMLError as e:
        raise ValueError(f"Invalid YAML in {path}: {e}")

    # Add this
    if data is None:
        return {}

    if not isinstance(data, dict):
        raise ValueError(f"Config must be a mapping, got {type(data)}")

    return data


# region Temporary Functions
def normalize_yaml_for_dataclass(yaml_data: dict) -> dict:
    """Convert YAML strings to proper types for UserConfig"""
    normalized = yaml_data.copy()

    # Convert chunk_type to ChunkType enum
    if "chunk_type" in normalized and isinstance(normalized["chunk_type"], str):
        try:
            normalized["chunk_type"] = ChunkType(normalized["chunk_type"])
        except ValueError as e:
            raise ValueError(
                f"{e}"
                f"Invalid chunk_type: '{normalized['chunk_type']}'. "
                f"Valid options: {[c.value for c in ChunkType]}"
            )

    # Convert chunk_type to ChunkType enum
    if "direction" in normalized and isinstance(normalized["direction"], str):
        try:
            normalized["direction"] = PipelineDirection(normalized["direction"])
        except ValueError as e:
            raise ValueError(
                f"{e}"
                f"Invalid chunk_type: '{normalized['direction']}'. "
                f"Valid options: {[d.value for d in PipelineDirection]}"
            )

    return normalized


# endregion
