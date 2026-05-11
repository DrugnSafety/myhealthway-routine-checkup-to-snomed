# Changelog

All notable changes to this project.

## [0.1.0] — 2026-05-06 (Cowork plugin packaging)

### Added
- Packaged as Cowork `.plugin` with manifest, slash commands, MCP server, and embedded skill.
- 4 slash commands: `/convert-checkup`, `/batch-convert-checkup`, `/verify-v5-workbook`, `/compare-exam-dates`.
- 5 MCP tools (FastMCP server): `convert_subject`, `batch_convert`, `verify_workbook_v5`, `get_codes`, `normalize_subject`.
- Golden reference workbook (`신정순_v5_canonical.xlsx`) for regression testing.
- Anonymization helper (output filenames use 초성 + 생년월일).

## [0.0.5] — 2026-05-06 (v5 format, Python deterministic core)

### Added
- v5 5-sheet workbook spec implemented in Python (`v5_core.py` + `v5_writer.py`).
- `convert_to_v5.py` (single subject), `batch_convert.py` (77 patients), `verify_workbook.py` (schema check + golden compare).
- 18-column schema for `A_hc_finding` (raw measurement findings) and `B_hc_clinical` (clinical interpretation).
- 5대 원칙: ①양성소견 present 직접표기 ②상위진단확정 시 하위 생략 ③음성확인이 임상적 의미 있을 때만 absent ④not_measured=비해당 ⑤always-generate 13항목.
- presence_status 결정 트리(`RULES.md`)·SNOMED+LOINC 매핑(`CODES.json`)·예시 문서.
- 93.7% match rate with hand-crafted canonical reference; remaining ~6% is data not present in source JSON (stomach endoscopy findings).

### Changed
- LLM → 코드 책임 분리. 변환·검증은 결정론적 Python이 수행, LLM은 오케스트레이션만 담당.

## [0.0.3] — 2026-05-05 (Per-subject 5-sheet workbooks)

### Added
- Per-subject Excel files with 5 sheets: 환자정보 / hc_main / hc_finding / cancer_screening / lifestyle.
- 77 unique patients converted (83 JSON entries; 6 dedup'd).
- 색인 워크북 (index of all subjects).
- SNOMED CT MCP-verified codes (Hypertension 38341003, Diabetes 73211009, Anemia 271737000, Central obesity 248311001, Hyperlipidemia 55822004, etc.).

## [0.0.2] — 2026-05-05 (Multi-patient flat Excel)

### Added
- `건강검진데이터_processed_v3.xlsx`: 83 patients × 132 columns flat schema.
- ■/□ lifestyle parsing logic.
- Cancer screening arrays unfolded into indexed columns.

## [0.0.1] — 2026-05-05 (Initial)

### Added
- FHIR JSON parsing for `routineCheckupDataResponse`.
- 132-column schema specification.

[0.1.0]: https://github.com/DrugnSafety/myhealthway-routine-checkup-to-snomed/releases/tag/v0.1.0
[0.0.5]: https://github.com/DrugnSafety/myhealthway-routine-checkup-to-snomed/releases/tag/v0.0.5
[0.0.3]: https://github.com/DrugnSafety/myhealthway-routine-checkup-to-snomed/releases/tag/v0.0.3
[0.0.2]: https://github.com/DrugnSafety/myhealthway-routine-checkup-to-snomed/releases/tag/v0.0.2
[0.0.1]: https://github.com/DrugnSafety/myhealthway-routine-checkup-to-snomed/releases/tag/v0.0.1
