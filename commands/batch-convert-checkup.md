---
description: "여러 환자 건강검진 데이터(JSON 리스트 또는 132-컬럼 엑셀)를 한 번에 v5 워크북으로 일괄 변환"
argument-hint: "[입력파일] [출력디렉터리]"
allowed-tools: Bash, Read, Write
---

You are running the `/batch-convert-checkup` command. The user wants to convert all subjects in an input source into per-patient v5 SNOMED CT workbooks.

# Arguments

User-provided arguments: `$ARGUMENTS`

The arguments should contain:

- An input file path: either a JSON list of envelopes (preferred) or a 132-column flattened xlsx.
- An optional output directory. If absent, default to `<workspace>/대상자별_v5/`.

If the user did not specify either, check the most recent attached file. If still ambiguous, ask in one short message.

# Workflow

1. **Identify input type**:
   - `.json` → use `--json` flag.
   - `.xlsx` → use `--xlsx` flag.
   - Other → reject and ask user to provide JSON or xlsx.

2. **Decide output directory**:
   - User-provided > default `${workspace}/대상자별_v5/`.
   - Ensure it exists.

3. **Run batch conversion**:
   ```bash
   cd "${CLAUDE_PLUGIN_ROOT}/skills/snomed-ct-mapper/scripts"
   python3 batch_convert.py \
     --json "<input>" \  # or --xlsx
     --out-dir "<out_dir>" \
     --index "<out_dir>/_INDEX_v5.xlsx"
   ```

4. **Surface progress and summary**:
   - The script prints one line per subject (`[N/total] OK ...`) and a final JSON summary.
   - Show the `total / ok / failed` counts to the user.
   - If any failed, list them with their error messages.

5. **Sample verification**:
   - Run `verify_workbook.py` on 2–3 random output files (first, middle, last) to confirm no systemic regression.
   - If all pass, just report PASS. If any fail, surface details.

6. **Report to user**:
   - Total subjects converted (and how many had comparison mode active).
   - Output directory + index file as `computer://` link.
   - One sample patient file as `computer://` link.

# Failure handling

- If the script exits non-zero with a partial batch, surface the failed subjects and ask the user whether to retry only those.
- Never silently skip subjects.

# Output format to user

Keep concise:
1. Counts: total / ok / failed.
2. Output directory link + index link.
3. One-line note about comparison-mode subjects (if any).
4. Verification status of sampled files.
