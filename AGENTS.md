# XlsxPress agent instructions

Project-specific instructions. They take precedence over the global instructions where they conflict.

- This project follows the `python-dev`, `python-docstrings` and `python-unit-tests` skills.
- The design decisions (including all conversion decisions) are in `docs/design.md`. Keep it up to date when a decision changes, in the same commit.

## Data conversion is correctness-critical

Reading input files (`src/xlsxpress/core/readers.py`) and writing Excel files (`src/xlsxpress/core/writers.py`) decide whether values reach Excel unchanged. Users draw scientific conclusions from these workbooks.

- Test conversion code thoroughly: every data type, every Excel limit, and every decision in the decisions table of `docs/design.md`, including edge cases. This is the one place where exhaustive tests are wanted; elsewhere keep tests focused.
- Never change a value silently. A value is written unchanged, or the change is a documented decision, or the user gets a warning or an error.
- Check behaviour on the written `.xlsx` file itself (cell values, cell types, number formats), not only on the DataFrame.
