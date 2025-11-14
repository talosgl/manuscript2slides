# region imports
from __future__ import annotations


import logging
import os
from typing import Optional
from manuscript2slides.internals.config.define_config import UserConfig
from manuscript2slides.internals import constants
from manuscript2slides.utils import str_to_bool

log = logging.getLogger("manuscript2slides")

import argparse
import sys
from typing import Optional
from manuscript2slides.internals.constants import SENTINEL

# endregion


# region get_debug_mode
def get_debug_mode(
    cfg: UserConfig | None = None, interface_flag: bool | None = None
) -> bool:
    """
    Determine debug mode from multiple sources.
    Priority: just-passed interface arg > existing in-use config (from file, cli, gui) > env var > DEBUG_MODE_DEFAULT (constants.py)

    Priority order (highest to lowest):
    1. just-passed interface arg (interface_flag) - probably from a gui click
    2. CLI argument (--debug true/false) (checked via check_for_cli_debug_arg())
    3. Config object value (cfg.debug_mode, from file, gui, or argparse)
    4. GUI preference (QSettings)
    5. Environment variable (MANUSCRIPT2SLIDES_DEBUG)
    6. internals/constants.py DEBUG_MODE_DEFAULT

    Args, both optional and will be read as None if not provided:
        cfg: Configuration object (may contain debug_mode setting)
        interface_flag: Explicit CLI arg or GUI arg value passed in

    Returns:
        bool: Whether debug mode should be enabled from this point on
    """

    # 1. Highest Priority: ad-hoc CLI/GUI argument (explicit user override)
    if interface_flag is not None:
        return interface_flag

    cli_flag = _check_for_cli_debug_arg()
    # 2. Check cli args
    if cli_flag is not None:
        return cli_flag

    # 3. Second Priority: Config object value (built from gui, loaded from file, built from argparse)
    if cfg is not None and cfg.debug_mode is not None:
        return cfg.debug_mode

    # 4. GUI preference (QSettings)
    from manuscript2slides.gui import APP_SETTINGS

    gui_pref = APP_SETTINGS.value("debug_mode")
    if gui_pref is not None:
        # QSettings returns strings, need to convert
        return str_to_bool(gui_pref.lower())

    # 5. Check env variable
    env_debug_str = os.environ.get("MANUSCRIPT2SLIDES_DEBUG")
    if env_debug_str is not None:
        try:
            # If a valid value is found, return it immediately
            return str_to_bool(env_debug_str)
        except ValueError:
            # If the env var is set but invalid ("bob"), log a warning and fall through to default
            print(
                f"Warning: Invalid value for MANUSCRIPT2SLIDES_DEBUG env var: '{env_debug_str}'. Using default."
            )

    # 6. Lowest Priority / Fallback: The system default constant
    return constants.DEBUG_MODE_DEFAULT


# endregion


# region _check_for_cli_debug_arg
def _check_for_cli_debug_arg() -> Optional[bool]:
    """
    Uses a dedicated ArgumentParser instance to safely extract only the debug flag value.
    Returns the value (True/False) or None if the flag was the SENTINEL/not provided.
    """

    # Create a micro parser just for this purpose
    parser = argparse.ArgumentParser(add_help=False)  # Turn off help
    parser.add_argument(
        "--debug",
        "-dbg",
        "--dbg",
        "--debug-mode",
        "--debug_mode",
        dest="debug_mode",
        type=str_to_bool,
        metavar="BOOL",
        default=SENTINEL,
    )

    # Parse ONLY the known arguments, ignoring everything else
    # This prevents errors if other, unrelated args are present
    args, _ = parser.parse_known_args(sys.argv)

    if args.debug_mode is not SENTINEL:
        # If the sentinel was overwritten, the user provided a value
        return args.debug_mode
    else:
        # The user did not provide the flag
        return None


# endregion
