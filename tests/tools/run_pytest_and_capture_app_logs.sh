#!/bin/bash
# Run pytest with verbose output and capture to .logs/pytest_output.log
#
# This script is useful for reviewing test output after the fact, especially
# when debugging test failures or examining application logs generated during tests.
#
# Usage (from project root):
#   ./tests/tools/run_pytest_and_capture_app_logs.sh
#
# Output:
#   - Console: Live pytest output (verbose mode)
#   - File: .logs/pytest_output.log (appended, not overwritten)

mkdir -p .logs
pytest -v | tee -a .logs/pytest_output.log