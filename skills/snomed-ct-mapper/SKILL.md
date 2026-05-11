---
name: snomed-ct-mapper
description: "한국 국가건강검진(나의건강기록 앱 JSON 또는 평탄화된 132-컬럼 엑셀) 데이터를 SNOMED CT 표준코드로 매핑하여 v5 포맷(5-sheet 워크북)으로 출력. 트리거: 'SNOMED 매핑', 'hc_finding 생성', 'v5 포맷', '검진결과 분석', '건강검진 SNOMED', '대상자별 분석', '임상소견 코드화' 등."
version: 5.1
last_updated: 2026-05-06
---

# snomed-ct-mapper — v5 포맷 (Korean Health Checkup → SNOMED CT)

## 동작 방식 (Important)

이 skill은 **결정론적 Python 스크립트**가 변환을 수행하고, LLM은 **오케스트레이션·검증·보고**만 담당합니다. 18-컬럼 스키마, 5대 원칙, always-generate 13항목, ■/□ 파싱, 단위 변환은 모두 코드(`scripts/v5_core.py`, `v5_writer.py`)에 박혀 있습니다.

LLM이 수행할 작업: ① 입력 파일 식별 → ② 적절한 스크립트 호출 → ③ 결과 검증(`verify_workbook.py`) → ④ 사용자에게 결과 보고.

LLM이 수행하지 말 것: 워크북 내용을 직접 텍스트로 생성하거나, SNOMED 코드를 즉석에서 매핑하거나, 18-컬럼 행을 LLM이 짜내는 행위. 모든 데이터 변환은 스크립트가 담당해야 합니다.

---

## 입력

- **JSON 입력** — `routineCheckupDataResponse` (단건/배열). 같은 환자의 두 검진일자 발견 시 자동 비교 모드.
- **엑셀 입력** — 132-컬럼 평탄화 시트 (이전 v3 산출물).

---

## 호출 명령

### 단일 환자 변환

```bash
cd "<workspace>/skills/snomed-ct-mapper/scripts"
python3 convert_to_v5.py --json <INPUT.json> --name "<환자명>" --out <OUT.xlsx>
# 또는 엑셀 행 지정
python3 convert_to_v5.py --xlsx <FLAT132.xlsx> --row 26 --out <OUT.xlsx>
# 두 일자 비교 모드
python3 convert_to_v5.py --json <CURR.json> --name "<환자명>" --compare-with <PREV.json> --out <OUT.xlsx>
```

stdout에 JSON 보고서를 출력합니다 — `{status, output, patient, patient_id, exam_date, a_rows, b_rows, c_rows, comparison}`.

### 일괄 변환 (전체 데이터셋)

```bash
python3 batch_convert.py --json <LIST.json> --out-dir <OUT_DIR> --index <INDEX.xlsx>
# 132-컬럼 엑셀 입력
python3 batch_convert.py --xlsx <FLAT132.xlsx> --out-dir <OUT_DIR> --index <INDEX.xlsx>
# 테스트 시 일부만
python3 batch_convert.py --json <LIST.json> --out-dir <OUT_DIR> --limit 5
```

같은 환자(name+birth)의 두 records가 발견되면 자동으로 비교 모드 적용. 처리 진행상황을 표준출력에 한 줄씩 표시.

### 검증

```bash
# 단독 검증 (스키마, presence_status, qualifier 정합성, always-generate 누락 체크)
python3 verify_workbook.py <OUT.xlsx>
# 골든 레퍼런스 대비 행 수 비교
python3 verify_workbook.py <OUT.xlsx> --golden <REF.xlsx>
# 행 수 차이를 에러로 처리(strict)
python3 verify_workbook.py <OUT.xlsx> --golden <REF.xlsx> --strict
```

종료 코드: `0`=PASS, `1`=경고만, `2`=에러. JSON 형식의 보고서 출력.

---

## 출력 워크북 구조 (v5)

| 시트 | 의미 | 핵심 |
|---|---|---|
| `00_환자기본정보` | 환자 식별·검진 메타·계측·혈액·요/영상 (5컬럼) | 두 검진일자 side-by-side |
| `A_hc_finding` | 측정값 직접 소견 (raw, 18컬럼) | A_계측~I_암검진 9개 카테고리 |
| `B_hc_clinical` | 임상 해석 (질환 present/absent, 18컬럼) | 5대 원칙 + 13개 always-generate |
| `C_2023_비교` | 비교 모드 시 차이 항목 (18컬럼) | ★YYYY 마커 |
| `D_SNOMED_코드_레퍼런스` | 신규/검증필요 코드 추적 (6컬럼) | ⚠ 마커 |

상세 18-컬럼 헤더, 카테고리 분리행, 5대 원칙 본문 등은 `references/TEMPLATE_v5.json` 참조.

---

## presence_status 결정 알고리즘 (요약)

코드에 박혀있는 결정 트리:

- 측정값 없음/비해당 → `not_measured`
- 정상 범위 → 질환 코드를 `absent`(refuted) (B시트에서 임상적 의미 있을 때만)
- 비정상 + 진단확정 단서(유질환자/치료중) → `present`(confirmed)
- 비정상 + 의심수준 단서 → `suspected`(unconfirmed)

`presence_qualifier_code`는 자동 매핑됨: `present`→`410515003`, `absent`→`410516002`, 그 외 `—`.

전체 결정 트리는 `references/RULES.md` 참조.

---

## "Always-generate" 항목 (B시트 13개)

검사 시행 여부와 무관하게 항상 행을 생성한다. 미시행/비해당이면 `not_measured`.

B형간염, 골다공증, 우울증, 인지기능장애, 낙상위험, ADL 의존, 요실금, 인플루엔자/폐렴구균 백신, 위/대장/유방/자궁경부암.

---

## 파일 트리

```
skills/snomed-ct-mapper/
├── SKILL.md                     ← 이 파일 (오케스트레이션 가이드)
├── scripts/
│   ├── v5_core.py               ← 데이터 레이어 (parsers, derivers)
│   ├── v5_writer.py             ← openpyxl 시트 직렬화
│   ├── convert_to_v5.py         ← 단일 환자 CLI
│   ├── batch_convert.py         ← 다중 환자 CLI
│   └── verify_workbook.py       ← 스키마·정합성 검증
└── references/
    ├── CODES.json               ← SNOMED + LOINC 매핑 (verified codes)
    ├── RULES.md                 ← presence_status 결정 트리, 단위 변환
    ├── TEMPLATE_v5.json         ← 워크북 골격 (시트명·헤더·스타일링)
    ├── EXAMPLE_신정순.md         ← 참조 예시
    └── golden/
        ├── 신정순_v5_canonical.xlsx   ← 사용자가 손수 만든 v5 reference
        └── 신정순_v5_generated.xlsx   ← 스크립트가 생성한 비교 대상
```

---

## 워크플로우 (LLM이 따를 단계)

1. **입력 인식** — 사용자 메시지 또는 첨부 파일에서 JSON/엑셀 식별. 단건인지 다건인지 확인.

2. **사용자 확인 (`AskUserQuestion`)** — 다음을 한 번에 묻는다:
   - 출력 모드: 단일 환자 1파일 / 환자별 분리 / 통합 long-format
   - 비교 모드 적용 여부 (두 검진일자 발견 시)
   - 출력 디렉터리

3. **스크립트 호출** — `convert_to_v5.py` (단일) 또는 `batch_convert.py` (다중).

4. **검증** — `verify_workbook.py`로 스키마/정합성 확인. 경고가 있으면 사용자에게 명시.

5. **사용자 워크스페이스에 복사** — `<workspace>/대상자별_v5/` 또는 사용자가 지정한 경로.

6. **사용자 보고** — `computer://` 링크와 함께 변환 결과 요약.

---

## 신규 SNOMED 코드 도입 시 (확장)

새로운 임상 개념이 등장하면:

1. SNOMED CT MCP로 코드 검증 (`mcp__snomed-ct__snomed_lookup`).
2. `references/CODES.json`에 코드 추가.
3. `v5_core.py`의 derivation 함수에 새 분기 추가.
4. `verify_workbook.py`의 `ALWAYS_GENERATE_KEYWORDS`에 추가 (always-generate 항목인 경우).
5. `references/golden/` 하의 reference 워크북과 회귀 테스트.

이 5단계를 거치면 변환의 안정성·재현성을 유지할 수 있습니다.

---

## 안정성 기준

- **정확성**: 같은 입력 → 같은 출력 (코드 기반 결정론적 변환)
- **검증**: `verify_workbook`이 5개 시트 존재, 18-컬럼 헤더, presence_status 정합성, always-generate 항목 누락 여부를 점검
- **회귀**: `references/golden/` 의 canonical 파일과 row-count diff로 추적
- **확장**: 새로운 코드는 `CODES.json`에 추가 → 즉시 모든 변환에 반영
