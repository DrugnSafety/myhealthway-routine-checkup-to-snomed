---
description: "생성된 v5 워크북이 스키마와 always-generate 규칙을 준수하는지 검증 (선택적으로 골든 레퍼런스와 비교)"
argument-hint: "[워크북파일] [--golden 골든파일]"
allowed-tools: Bash, Read
---

You are running the `/verify-v5-workbook` command. The user wants to validate one or more v5-format Excel workbooks against the v5 schema rules.

# Arguments

User-provided arguments: `$ARGUMENTS`

Parse from arguments:

- A target xlsx file path (required). If a directory is given, verify every `.xlsx` in it.
- An optional golden reference path (after `--golden` or as second xlsx file). Used for row-count comparison.
- Optional `--strict` flag: treat row-count divergence as error rather than warning.

# Workflow

## Single file mode

1. Run:
   ```bash
   cd "${CLAUDE_PLUGIN_ROOT}/skills/snomed-ct-mapper/scripts"
   python3 verify_workbook.py "<target>" [--golden "<ref>"] [--strict]
   ```

2. Parse the JSON output (the script prints a structured report).

3. Report to user:
   - **PASS** / **WARNINGS** / **ERRORS** status.
   - Sheet-by-sheet row counts (`stats`).
   - If golden was provided, diff vs golden.
   - If errors or warnings, list each with the sheet/row context.

## Directory mode

If the target is a directory, iterate over every `.xlsx` inside (excluding files starting with `_INDEX`):

1. Run `verify_workbook.py` on each.
2. Aggregate results: `n_pass`, `n_warn`, `n_err`.
3. Show a compact table: filename | status | A rows | B rows.
4. Surface any error files in detail.

# Failure handling

- If the script crashes (exit code other than 0/1/2), include stderr verbatim and stop.

# Output format to user

Keep concise:
- Status banner (PASS / WARNINGS / ERRORS).
- Summary stats (rows per sheet, or aggregated counts in directory mode).
- List of warnings/errors with sheet/row context (truncate if more than 10).
- Suggest fixes if errors are present (e.g., "missing always-generate item: 골다공증 → check derive_b_findings").
