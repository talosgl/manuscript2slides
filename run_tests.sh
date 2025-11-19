#!/bin/bash
mkdir -p .logs
pytest -v | tee -a .logs/pytest_output.log