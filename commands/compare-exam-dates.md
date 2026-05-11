---
description: "동일 환자의 두 검진일자 데이터를 비교하여 v5 워크북의 C 시트(차이 항목)를 채워서 생성"
argument-hint: "[환자명] [최신JSON] [이전JSON] [출력파일경로]"
allowed-tools: Bash, Read, Write
---

You are running the `/compare-exam-dates` command. The user wants a v5 workbook that explicitly compares two exam dates for the same patient — the C sheet will be populated with differences.

# Arguments

User-provided arguments: `$ARGUMENTS`

Parse from arguments (in any order):

- Patient name (Korean), e.g. 신정순.
- Two JSON file paths: the most recent exam JSON (current) and an older exam JSON (prior).
  - If only one JSON is given but the patient appears twice in it (different exam dates), use the same JSON for both `--json` and `--compare-with`.
- Optional output xlsx path. Default: `<workspace>/대상자별_v5/{patient}_{birth}_v5_compare.xlsx`.

# Workflow

1. **Validate inputs**:
   - Both files must exist and be parseable JSON.
   - The patient name must appear in both files.
   - If multiple records per file, the script picks by `examinationDate`; older becomes `prev`.

2. **Run the comparison conversion**:
   ```bash
   cd "${CLAUDE_PLUGIN_ROOT}/skills/snomed-ct-mapper/scripts"
   python3 convert_to_v5.py \
     --json "<current>" \
     --name "<patient>" \
     --compare-with "<prior>" \
     --out "<output>"
   ```

3. **Verify**:
   ```bash
   python3 verify_workbook.py "<output>"
   ```

4. **Highlight C sheet**:
   - Inspect the output: read the C sheet's rows and identify the most clinically significant changes.
   - Surface them as a short bullet list (3–5 items max). Examples: "혈압 159/78 → 146/74 (개선)", "ALT 98 → 24 (정상화)", "청력(좌) 질환의심 → 정상".

5. **Report**:
   - Patient + two exam dates.
   - C sheet row count and a 3–5 item highlight summary.
   - `computer://` link to the output file.

# Failure handling

- If only one date exists for the patient, fall back to single-mode `/convert-checkup` and tell the user the comparison was not possible.
- If the script reports `c_rows: 0`, double-check the inputs and warn the user that the two records may be identical or share the same date.

# Output format to user

Keep concise:
1. Patient + dates compared.
2. Number of differences detected.
3. 3–5 most significant changes (one line each).
4. `computer://` link.
