"""Test suite for manuscript2slides.

This package contains all automated tests for the manuscript2slides application,
organized to mirror the source code structure.

Running Tests:
    pytest                                  # Run all tests
    pytest -v                               # Verbose output
    pytest tests/test_cli.py                # Run specific file
    pytest -s                               # Don't capture output (for debugging)
    pytest -k "test_boolean"                # Run tests with matching pattern in function name

Capture/append pytest stdout output to file:
    run_tests.sh # script containing the below bash lines

    # or
    mkdir -p .logs
    pytest -v | tee -a .logs/pytest_output.log

Debugging Tests:
    - Use breakpoint() in test code, then run with pytest -s
    - Set breakpoints in VS Code and use "Debug Current Test" launch config
    - Use pytest --pdb to drop into debugger on failure
    - For gdb-like watchpoints, `pip install watchpoints`, then and use
        `watch(variable_name)` in test code, then run with pytest -s

Coverage:
    pytest --cov=manuscript2slides --cov-report=html
    # Then open htmlcov/index.html

Notes:
    - Monkeypatch for changing values (sys.argv, env vars)
    - Mock/patch for spying on function calls and faking behavior
    - Aim for testing behavior, not implementation details
"""
