---
name: python-docstrings
description: Format and write Google-style docstrings for Python code compatible with mkdocstrings. Use whenever writing, generating, documenting, or modifying Python functions, classes, or modules.
metadata:
  version: 2026-11-06
---

# Python Docstring Guidelines

Generate Google-style docstrings formatted for compatibility with `mkdocstrings`. Write in American English using ASCII characters only. The target audience includes domain scientists (proteomics, bioinformatics) and general software developers.

## Usage Policy for Examples
- **Do NOT** include examples for trivial or self-explanatory functions.
- **ONLY** include an `Examples:` section for functions where extra context is useful, and for primary high-level public API functions.

## Formatting Rules & Code References
- Use single backticks (`) for all code elements, types (e.g., `None`), and string literals (e.g., `proquant-notebook`).
- Try to avoid nesting single quotes inside backticks unless necessary to clarify meaning (avoid `'text'`).
- For explicit Python syntax or parameter values, use assignment format (e.g., `style="proquant-notebook"`).

## Section Structure & Order
1. **Title**: Imperative summary line.
2. **General explanations**: Paragraph detailing functionality or context if not obvious from the title.
3. **Args**: Parameter descriptions formatted per default-value rules.
4. **Returns** (or **Yields** for generators): Return type and description.
5. **Raises**: Exceptions raised and their triggers.
6. **Examples** *(Conditional)*:
  - Use `>>>` doctest style **only** when the output is predictable and the snippet can actually be executed (i.e., a doctest would pass). Do not use "imagined" or mock inputs that cannot run.
  - Use a fenced code block otherwise.

## Default Value Description Rules in `Args`
- **For `bool`**: Add a `(default)` tag directly in the action statement:
  - Example: `If True (default), normalizes data before returning.`
- **For `None` with fallback behavior**: Use a conditional sentence:
  - Example: `If None (default), calculates mean values dynamically.`
- **For `None` as optional/omitted**: Use a trailing sentence:
  - Example: `Optional output directory path. Defaults to None.`
- **For specific scalar values**: Use an explicit trailing sentence or assignment style:
  - Example: `Maximum retry attempts for requests. Defaults to 3.`