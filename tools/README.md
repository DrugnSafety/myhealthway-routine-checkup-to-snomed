# Anonymization tools

## Usage

```bash
# v5 per-subject files (e.g. 신정순_1949-05-30_v5.xlsx)
python3 tools/anonymize_by_choseong.py \
    /path/to/per-subject-output \
    /path/to/anonymized-output \
    /path/to/private/mapping.csv

# Older flat 5-sheet files (e.g. 신정순_1949-05-30.xlsx)
python3 tools/anonymize_v3.py \
    /path/to/source \
    /path/to/anonymized \
    /path/to/mapping.csv
```

Output filenames follow the pattern: `routineCheckup_{초성}_{YYYY-MM-DD}.xlsx`. Collisions are disambiguated with a 4-char SHA-256 hash.

The mapping CSV is **never** committed — it links real names ↔ anonymized identifiers.
