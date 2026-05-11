# v5 Mapping Rules — Detailed Reference

## 1. presence_status 결정 트리

```
입력: 측정값 v, 정상범위 R, 사용자 진술 U (e.g. "유질환자", "치료중", "비해당", "의심")

if v is None or v == '비해당' or v == '미시행':
    → not_measured        (qualifier_code = '—')
elif v ∈ R:                 # 정상 범위 내
    A시트:  '정상' 소견 코드를 present
    B시트:  해당 질환을 absent (refuted) — 음성확인이 의미있을 때만
elif v ∉ R and U ∈ {"유질환자","치료중","약물복용중"}:
    → present + verification_status = confirmed
elif v ∉ R and U == "고혈압-전단계":
    → present + verification_status = differential (전단계 코드 사용)
elif v ∉ R and U ∈ {"의심","질환의심","경계","공복혈당장애 의심"}:
    → suspected + verification_status = unconfirmed
else (v ∉ R, no clinical context):
    → present + verification_status = unconfirmed (보수적)
```

## 2. presence_qualifier_code 매핑

| presence_status | qualifier_code | 비고 |
|---|---|---|
| present         | 410515003 (Known present)         | known_treated 도 동일 |
| absent          | 410516002 (Known absent)          | refuted 와 짝 |
| suspected       | —                                 | qualifier 없음 |
| not_measured    | —                                 | qualifier 없음 |

## 3. verification_status 의미

- `confirmed` — 진단/측정값 + 임상 진술이 모두 일치 (예: 혈압 146 + 유질환자)
- `refuted`   — 음성/정상으로 확인된 absent (예: 청력 정상 → 난청 absent)
- `unconfirmed` — 측정값만 존재하고 추가 확인 없음 (의심 단계)
- `differential` — 상위/하위 진단 후보 중 선택 표기 (예: prehypertension)

## 4. 라이프스타일 ■/□ 해석

`examination.lifestyleFactors` 원문:
```
■ 금연 필요 □ 절주 필요 ■ 신체활동 필요 ■ 근력운동 필요
```

- `■ 금연 필요` → 환자가 흡연중 → A시트에 77176002 Smoker present, B시트에도 동일
- `□ 금연 필요` → 환자가 비흡연 → A시트에 266919005 Never smoked tobacco present, B시트에 77176002 Smoker absent
- 절주: ■ → 위험음주 present (228273003 또는 160592001)
- 신체활동: ■ → 43994002 Lack of physical activity present
- 근력운동: ■ → 40979000 Lack of exercise present

## 5. always-generate (B시트) — 13개 항목

검사 시행 여부와 무관하게 B시트에 행을 생성한다. 미시행/비해당이면 `not_measured`.

| 항목 | SNOMED |
|---|---|
| B형간염 | 66071002 |
| 골다공증 | 64859006 |
| 우울증 | 35489007 |
| 인지기능장애 | 386806002 |
| 낙상 위험 | 129839007 |
| ADL 의존 | 160674001 |
| 요실금 | 165232002 |
| 인플루엔자 백신 | 86198006 |
| 폐렴구균 백신 | 12866006 |
| 위암 | 363349007 |
| 대장암 | 363406005 |
| 유방암 | 254837009 |
| 자궁경부암 | 363354003 |

## 6. 단위 변환 규칙

| 원시값 | 변환 | 예시 |
|---|---|---|
| weight 4자리 정수 | ÷10 | "0873" → 87.3 kg |
| height 4자리 정수 | ÷10 | "1576" → 157.6 cm |
| egfr 4자리 정수 | 첫 0 제거 | "0097" → 97 |
| serumCreatinine 정수(0.x로 추정) | ÷10 | "6" → 0.6 (단 "0.78"은 그대로) |
| visionLeft/Right 2자리 정수 | ÷10 | "10" → 1.0, "08" → 0.8 |
| fastingGlucose 3자리 정수 | 그대로 | "163" → 163 |

원시값과 해석값을 모두 보존: 00시트 비고 또는 mapper_note에 "원시:6 → 0.6 해석"

## 7. 상위진단 우선 원칙 (B시트 ②)

다음 쌍에서 상위 진단이 confirmed이면 하위 행을 생략:

| 상위 (생략 트리거) | 하위 (생략 대상) |
|---|---|
| Hypertension confirmed | Prehypertension, Elevated BP reading |
| Diabetes mellitus confirmed | Impaired fasting glycaemia, Hyperglycemia |
| Obesity confirmed | Overweight |

다만 BMI = 26.7 처럼 **상위가 아직 미달**하면 두 코드 모두 표기 (Overweight present + Obesity absent).

## 8. 비교 모드 차이 추출 알고리즘

```
for item in (A∪B의 모든 finding):
    v1 = item.exam_date_1.derived_value
    v2 = item.exam_date_2.derived_value
    s1 = item.exam_date_1.presence_status
    s2 = item.exam_date_2.presence_status
    if s1 != s2:                                       # 상태 변화
        → C시트에 추가 (★YYYY 마커)
    elif s1 == s2 == 'present' and significantly_different(v1, v2):
        → C시트에 추가 (수치 변화)
    elif one is not_measured and the other is present:
        → C시트에 추가 (검사 추가/누락)
```

## 9. SNOMED 코드 결정 우선순위

1. 본 SKILL의 references/CODES.json에 매핑이 있으면 그 코드 사용
2. 없으면 SNOMED CT MCP (`mcp__snomed-ct__snomed_lookup` 또는 `snomed_get_by_code`)로 검증
3. MCP 결과가 references/CODES.json과 다르면 **MCP 결과를 우선**하고 CODES.json 갱신
4. MCP가 응답하지 않으면 본 SKILL의 코드를 잠정 사용하고 D시트에 ⚠ 표기

## 10. mapping_confidence 등급

- `high` — 표준 코드, MCP로 확인됨, 기존 검증된 매핑
- `medium` — 합리적이나 부분 일치, 더 구체적 코드가 있을 수 있음
- `low` — 임시 매핑, 검증 필요, ⚠ 마커와 함께 D시트에 등록
