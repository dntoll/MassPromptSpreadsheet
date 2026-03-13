# Test CSV for PROMPT

The file `spreadsheet-prompt.csv` in this folder contains rows that test `=PROMPT(...)` in various ways.

## Import into Google Sheets

1. Create a new spreadsheet or open one where Mass Prompt is installed.
2. **File → Import → Upload** (or paste): select `spreadsheet-prompt.csv` from the `test/` folder.
3. On import: choose to replace or insert at current sheet; delimiter **comma**.
4. Set the API key via **Mass Prompt → Set API key** if you haven’t already.

## Contents

- **Column A:** Description of the test case.
- **B–D:** Inputs 1–3 (values used in the prompt).
- **E:** Prompt template (with `{0}`, `{1}`, etc., and optionally `[field1,field2]` for multiple outputs).
- **F:** Formula that calls PROMPT with the cells on the same row.

Rows that test errors require you to trigger them (e.g. clear the API key for the "Missing API key" row).

## Test cases

| Type | Description |
|------|-------------|
| Error | Too few arguments (prompt only) → `ERROR: At least one input and one prompt required` |
| Single output | 1, 2, or 3 inputs with a simple prompt |
| Multiple outputs | Prompt with `[firstname,lastname]` or `[country,city]` → spill to the right |
| Single field in [] | `[city]` → one extra column (spill) |
| Empty input | Prompt with empty `{0}` |
| URL input | Public webpage/PDF URL in input cell is fetched and used as `{0}` context |
| Chained dependency | Two-row chain where downstream input references upstream PROMPT output (e.g. `B16 = F15`) |
| Manual rerun token | Use `MP_REFRESH(token)` as first arg; change token to rerun one cell only |
| Error | Missing API key → `ERROR: Set API key via menu...` (clear key and reload) |
| Error | Multi-output with non-JSON response → `ERROR: Invalid structured response` (depends on model output) |
