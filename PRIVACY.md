# Privacy & Anonymization

This repository follows a privacy-by-default policy for patient data.

## What is committed
- Plugin code (skills, scripts, MCP server, commands)
- Anonymized example workbooks where patient names are replaced with 초성 (Korean consonant initials) + 생년월일

## What is NEVER committed
- Real patient names (Hangul syllable names)
- Mapping CSVs that link real names to anonymized identifiers (kept locally only)
- Raw FHIR JSON exports from 나의건강기록 app
- Any file inside `_PRIVATE_DO_NOT_COMMIT/` or matching `*_mapping.csv`

## 초성 anonymization scheme
For each Korean syllable, only the leading consonant (choseong) is kept:
- 신정순 (1949-05-30) → `ㅅㅈㅅ_1949-05-30`
- 안혜영 (1988-09-10) → `ㅇㅎㅇ_1988-09-10`

If two distinct patients hash to the same `(choseong, birth)` tuple, a 4-character SHA-256 hash of the original name is appended for disambiguation.

## Workflow

When generating workbooks locally:
1. Process the real JSON through `convert_to_v5.py` (outputs use real names).
2. Run the anonymization script to rename files for the repo.
3. Keep the mapping CSV in `_PRIVATE_DO_NOT_COMMIT/` (which is gitignored).
4. Commit only the anonymized files.

Mapping CSVs allow the local researcher to reverse-lookup if needed, while the repo itself contains no PHI/PII.
