# myhealthway-routine-checkup-to-snomed

한국 국가건강검진 데이터(나의건강기록 앱 JSON 또는 평탄화된 132-컬럼 엑셀)를 **SNOMED CT 표준코드** 기반 **v5 5-시트 워크북**으로 결정론적 변환하는 Cowork 플러그인입니다.

## 무엇을 하나요?

- 나의건강기록 앱에서 추출한 `routineCheckupDataResponse` JSON을 입력받아
- 18-컬럼 long-format 시트(`A_hc_finding`, `B_hc_clinical`)에 SNOMED CT 코드로 매핑하고
- 환자기본정보(`00`), 두 검진일자 비교(`C`), 코드 레퍼런스(`D`) 시트를 함께 생성합니다.
- 변환은 결정론적 Python 스크립트가 수행하므로 같은 입력에 대해 항상 같은 출력이 보장됩니다.

## 사전 요구

- Python 3.10+
- `openpyxl` 패키지: `pip install openpyxl`
- (선택) MCP 서버 사용 시: `pip install "mcp[cli]"`

## 사용 가능한 슬래시 명령

| 명령 | 용도 |
|---|---|
| `/convert-checkup` | 단일 환자 JSON → v5 워크북 |
| `/batch-convert-checkup` | 여러 환자 JSON 또는 132-컬럼 엑셀 → 환자별 v5 파일들 |
| `/verify-v5-workbook` | 생성된 워크북의 스키마·always-generate 항목 검증 (선택적 골든 비교) |
| `/compare-exam-dates` | 같은 환자 두 검진일자 비교 → C 시트 채워서 생성 |

## MCP 도구 (다른 세션에서도 호출 가능)

이 플러그인이 활성화되면 다음 MCP 도구가 자동으로 노출됩니다:

| 도구 | 인자 | 반환 |
|---|---|---|
| `convert_subject` | input_json_path, output_xlsx_path, patient_name?, compare_with_json_path? | 변환 보고서 |
| `batch_convert` | input_path, output_dir, input_kind('json'/'xlsx'), limit? | 일괄 변환 보고서 |
| `verify_workbook_v5` | workbook_path, golden_path?, strict? | 검증 결과(errors/warnings/stats) |
| `get_codes` | section? | SNOMED+LOINC 코드 매핑 테이블 |
| `normalize_subject` | input_json_path, patient_name? | 단위 변환된 평탄화 환자 dict (검사 시행 전 inspection용) |

## 출력 워크북 구조 (v5)

| 시트 | 컬럼 수 | 의미 |
|---|---|---|
| `00_환자기본정보` | 5 | 환자 식별·검진 메타·계측·혈액·요/영상 — 두 검진일자 side-by-side |
| `A_hc_finding` | 18 | 측정값 직접 소견 (raw) — A_계측 ~ I_암검진 9개 카테고리 |
| `B_hc_clinical` | 18 | 임상 해석 (질환 present/absent) — 5대 원칙 + 13개 always-generate |
| `C_2023_비교` | 18 | 비교 모드 시 차이 항목 — `★YYYY` 마커로 어느 일자가 다른지 표시 |
| `D_SNOMED_코드_레퍼런스` | 6 | 신규/검증필요 코드 추적 — `⚠` 마커 |

## A vs B 시트의 정신 차이

같은 측정값이 두 시트에 다르게 표현됩니다.

| 입력 | A 시트 (raw) | B 시트 (clinical) |
|---|---|---|
| 혈압 146/74 + 유질환자 | `24184005 Elevated BP reading` present | `38341003 Hypertension` present (confirmed) |
| 공복혈당 163 | `80394007 Hyperglycemia` present | `73211009 Diabetes mellitus` suspected |
| BMI 26.7 | `238131007 Overweight` present | Overweight present + `414916001 Obesity` absent |
| 청력 양이 정상 | `162339002 Hearing normal` present | `15188001 Hearing loss` absent (refuted) |

A는 측정사실, B는 질환 판정.

## "Always-generate" 13항목 (B 시트)

검사 시행 여부와 무관하게 행을 항상 생성합니다 (미시행 시 `not_measured`):

B형간염, 골다공증, 우울증, 인지기능장애, 낙상위험, ADL 의존, 요실금, 인플루엔자/폐렴구균 백신, 위/대장/유방/자궁경부암.

## 빠른 시작

1. 이 플러그인을 Cowork에 설치합니다.
2. 나의건강기록 JSON 파일을 채팅에 첨부합니다.
3. `/convert-checkup` 또는 `/batch-convert-checkup` 명령을 실행합니다.
4. 생성된 워크북은 `<workspace>/대상자별_v5/` 디렉터리에 저장됩니다.

## 검증·회귀 테스트

플러그인은 `skills/snomed-ct-mapper/references/golden/` 에 골든 레퍼런스 워크북(`신정순_v5_canonical.xlsx`)을 포함합니다. 변환 로직 변경 후 `/verify-v5-workbook` 명령에 `--golden` 옵션을 주면 row-count 회귀를 자동 점검합니다.

## 확장 — 새 코드 추가 시

1. SNOMED CT MCP로 코드 검증 (`mcp__snomed-ct__snomed_lookup`).
2. `skills/snomed-ct-mapper/references/CODES.json`에 코드 추가.
3. `scripts/v5_core.py`의 derivation 함수에 새 분기 추가.
4. `scripts/verify_workbook.py`의 `ALWAYS_GENERATE_KEYWORDS`에 추가 (해당하는 경우).
5. 골든 레퍼런스로 회귀 테스트.

## 라이선스

MIT
