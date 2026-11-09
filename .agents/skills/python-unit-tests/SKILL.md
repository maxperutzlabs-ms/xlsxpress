---
name: python-unit-tests
description: Generate and write pytest unit and integration tests. Use whenever writing, modifying, running, or refactoring Python unit or integration tests.
metadata:
  version: 2026-11-06
---

# Unit & Integration Testing Guidelines

Follow these standards when creating or updating tests for this project.

## Framework & Environment
- Use `pytest` for all unit and integration tests.
- Line length limit for test files is strictly **120 characters**. 

## Directory & File Layout
- **Unit tests**: Place inside `tests/unit/`.
- **Integration tests**: Place inside `tests/integration/`.

## Structure & Naming Conventions
- Group tests into classes named after the function or class being tested (e.g., `class TestFunction:`).
- Use class methods for individual test cases.
- Method names must be descriptive and self-documenting (e.g., `test_process_peptides_raises_value_error_on_invalid_scale`).
- Do **NOT** add docstrings or inline comments to test methods; the method name must completely explain the test intent and scenario.
