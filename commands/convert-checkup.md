---
description: "단일 환자 건강검진 JSON을 v5 5-시트 SNOMED CT 워크북으로 변환"
argument-hint: "[환자명] [입력파일경로] [출력파일경로]"
allowed-tools: Bash, Read, Write
---

You are running the `/convert-checkup` command for the snomed-ct-mapper plugin.

The user wants to convert a single patient's Korean national health checkup data (나의건강기록 JSON) into a v5-format SNOMED CT mapped 5-sheet Excel workbook.

# Arguments

User-provided arguments: `$ARGUMENTS`

The arguments may contain (in any order, separated by spaces or commas):

- A patient name in Korean (e.g., 신정순, 안혜영). Optional — if the input JSON contains only one subject, this is unnecessary.
- An input JSON file path. Required. May be the user's attached upload, a path in the workspace, or `/Users/.../uploads/...`.
- An output xlsx file path. Optional — if missing, default to `<workspace>/대상자별_v5/{patient_name}_{birth}_v5.xlsx`.

# Workflow

1. **Identify input file**:
   - If `$ARGUMENTS` contains a `.json` path, use it.
   - Otherwise check the most recent user-uploaded file or attached files in conversation.
   - If still ambiguous, ask the user (concise, one-line clarification).

2. **Identify patient name**:
   - If the user named a Korean patient, use that as `--name`.
   - Otherwise inspect the JSON: if it has only one entry, omit `--name`.
   - If multiple entries and no name, list them and ask which one.

3. **Decide output path**:
   - User-provided > default `${workspace}/대상자별_v5/{patient}_{birth}_v5.xlsx`.
   - Create the parent directory if it does not exist.

4. **Run the conversion**:
   ```bash
   cd "${CLAUDE_PLUGIN_ROOT}/skills/snomed-ct-mapper/scripts"
   python3 convert_to_v5.py --json "<input>" --name "<patient>" --out "<output>"
   ```
   Capture stdout (JSON report) and present `a_rows / b_rows / c_rows` counts to the user.

5. **Verify**:
   ```bash
   python3 verify_workbook.py "<output>"
   ```
   Show pass/fail status. If errors, surface them and offer to investigate.

6. **Report**:
   - Brief summary: patient, exam date, sheet row counts, verification status.
   - Provide a `computer://` link to the output file.
   - If `c_rows > 0`, note that comparison mode was active.
   - If `mapping_confidence` warnings exist or `verify_workbook` flagged missing always-generate items, surface them.

# Failure handling

- If the script exits non-zero, print the stderr/stdout and stop. Do not attempt to recreate the workbook by hand.
- If the input JSON is malformed, ask the user to verify the file rather than guessing.
- Never bypass the script and write the workbook directly with openpyxl — the SKILL document explicitly forbids that.

# Output format to user

Keep response concise:
1. One-line summary of what was converted.
2. Sheet row counts (00 / A / B / C / D).
3. Verification status (PASS / warnings / errors).
4. `computer://` link.
