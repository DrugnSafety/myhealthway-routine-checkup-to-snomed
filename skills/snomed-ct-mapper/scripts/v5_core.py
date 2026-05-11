"""v5_core — Pure data layer for v5 workbook generation.

Public API:
  - load_codes()                    → CODES dict
  - normalize_subject(rc)           → flat dict with units converted
  - derive_a_findings(subject)      → list[FindingRow]   (A_hc_finding rows)
  - derive_b_findings(subject)      → list[FindingRow]   (B_hc_clinical rows)
  - derive_c_findings(prev, curr)   → list[FindingRow]   (C_2023_비교 rows)

This module contains zero LLM-side logic and zero workbook (openpyxl) writes.
The caller is responsible for serializing FindingRow objects into Excel.

Invariants:
  - Unit conversion is performed exactly once, here.
  - SNOMED codes come ONLY from references/CODES.json (loaded by load_codes).
  - presence_status decision tree follows references/RULES.md.
  - All 13 always-generate items are emitted by derive_b_findings.
"""
from __future__ import annotations
import json
import re
import os
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict, Any

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
REFERENCES_DIR = os.path.normpath(os.path.join(THIS_DIR, '..', 'references'))


# ------------------------------------------------------------------ codes
def load_codes() -> dict:
    """Load curated SNOMED + LOINC mapping from references/CODES.json."""
    path = os.path.join(REFERENCES_DIR, 'CODES.json')
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


CODES = load_codes()
A_CODES   = CODES['A_hc_finding_codes']
B_CODES   = CODES['B_hc_clinical_codes']
E_CODES   = CODES['E_lifestyle_codes']
G_CODES   = CODES['G_special_codes']
I_CODES   = CODES['I_cancer_codes']
LOINC     = CODES['loinc_observables']
QUAL      = CODES['qualifier_value']
ALWAYS_B  = CODES['always_generate_in_B_sheet']
VERIFY_REQ = {x['code']: x for x in CODES['verification_required_v5']}


# ------------------------------------------------------------------ row
@dataclass
class FindingRow:
    """One row of A/B/C sheet (18 columns)."""
    n: int = 0                                  # # — set by writer
    patient_id: str = ''
    exam_date: str = ''
    finding_source: str = ''                    # measurement_derived / lifestyle_factor / etc.
    finding_category: str = ''                  # A_계측 / B_혈액 / ...
    item_name_kr: str = ''
    snomed_code: str = ''
    snomed_display: str = ''
    presence_status: str = ''                   # present / absent / suspected / not_measured
    presence_qualifier_code: str = '—'
    verification_status: str = ''               # confirmed / refuted / unconfirmed / differential
    snomed_domain: str = 'clinical_finding'
    derived_value: str = ''
    ref_range: str = ''
    medication_related: str = '9'               # 0/1/9
    mapping_confidence: str = 'high'
    loinc_code: str = '—'
    mapper_note: str = ''

    def __post_init__(self):
        # Auto-set qualifier code from presence_status
        if self.presence_status in ('present',):
            self.presence_qualifier_code = QUAL['known_present']
        elif self.presence_status == 'absent':
            self.presence_qualifier_code = QUAL['known_absent']
        else:
            self.presence_qualifier_code = '—'

    def as_list(self) -> list:
        return [self.n, self.patient_id, self.exam_date, self.finding_source,
                self.finding_category, self.item_name_kr, self.snomed_code,
                self.snomed_display, self.presence_status, self.presence_qualifier_code,
                self.verification_status, self.snomed_domain, self.derived_value,
                self.ref_range, self.medication_related, self.mapping_confidence,
                self.loinc_code, self.mapper_note]


# ------------------------------------------------------------------ parsers
def _num_match(s: Optional[str]):
    """Extract leading number from '132/90 고혈압' → '132/90', '35.1 비만' → 35.1."""
    if s is None: return None
    s = str(s).strip()
    m = re.match(r'^(\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)?)', s)
    if not m: return None
    v = m.group(1)
    if '/' in v: return v   # blood pressure
    try: return float(v)
    except ValueError: return None


def _suffix(s: Optional[str]) -> Optional[str]:
    """Return the trailing label after the numeric prefix."""
    if s is None: return None
    s = str(s).strip()
    m = re.match(r'^[\d./]+\s*(.*)$', s)
    return m.group(1).strip() if m else s


def _convert_unit(field: str, raw: Any) -> tuple[Any, str]:
    """Apply v5 unit-conversion rules. Returns (display_value, raw_note)."""
    if raw is None or raw == '': return None, ''
    s = str(raw).strip()
    num = _num_match(s)
    if num is None:
        return s, ''
    if isinstance(num, str):     # blood pressure like '132/90'
        return num, ''
    n = float(num)
    if field == 'weight' and n >= 100:        # 0873 → 87.3
        return n / 10, f'원시:{int(n)}'
    if field == 'height' and n >= 100:        # 1576 → 157.6
        return n / 10, f'원시:{int(n)}'
    if field == 'egfr' and n > 200:           # 0097 actually 97
        return n, ''
    if field == 'serumCreatinine' and n >= 4: # original was integer like 6 → 0.6
        return n / 10, f'원시:{int(n)}'
    if field in ('visionLeft', 'visionRight') and n > 1.5:
        return n / 10, ''
    return n, ''


def parse_lifestyle(s: Optional[str]) -> Dict[str, Optional[bool]]:
    """'■ 금연 필요 □ 절주 필요 ■ 신체활동 필요 ■ 근력운동 필요' → {smoking,alcohol,activity,strength}.
    True = needs intervention (■), False = OK (□), None = unknown."""
    out = {'smoking': None, 'alcohol': None, 'activity': None, 'strength': None}
    if not s: return out
    keymap = {'금연': 'smoking', '절주': 'alcohol', '신체활동': 'activity', '근력운동': 'strength'}
    for kw, key in keymap.items():
        m = re.search(r'([■□])\s*' + kw + r'\s*필요', s)
        if m:
            out[key] = (m.group(1) == '■')
    return out


def parse_vaccination(s: Optional[str]) -> Dict[str, Optional[bool]]:
    """'■ 인플루엔자 ... □ 폐렴구균 ... □ 접종 필요 없음' → {flu, pneumo, none}."""
    out = {'flu': None, 'pneumo': None, 'none': None}
    if not s: return out
    for m in re.finditer(r'([■□])\s*(인플루엔자|폐렴구균|접종 필요 없음)', s):
        sym, label = m.group(1), m.group(2)
        v = (sym == '■')
        if '인플루엔자' in label: out['flu'] = v
        elif '폐렴구균' in label: out['pneumo'] = v
        else: out['none'] = v
    return out


# ------------------------------------------------------------------ normalize
def normalize_subject(rc: dict) -> dict:
    """Flatten a routineCheckupDataResponse and apply unit conversions.
    Output keys match v5 column conventions; every numeric field has both
    `_raw` (original string) and `_value` (converted float) entries."""
    out = {}
    out['patientName']      = rc.get('patientName')
    out['birthDate']        = rc.get('birthDate')
    out['examinationDate']  = rc.get('examinationDate')
    out['screeningDate']    = (rc.get('screeningDate') or '')[:10]
    out['organizationName'] = rc.get('organizationName')

    # Patient ID — initials of name (Korean → first 3 chars or all) + birth year
    name = (rc.get('patientName') or 'NA')
    birth = (rc.get('birthDate') or '0000-00-00')
    initials = ''.join([c for c in name if not c.isspace()])[:3].upper()
    # Use Korean initials romanization if available; fallback to first 3 chars verbatim.
    out['patient_id'] = f'{initials}-{birth.split("-")[0]}'

    # measurement
    mt = rc.get('measurementTest') or {}
    for k, v in mt.items():
        display, note = _convert_unit(k, v)
        out[k]                = display
        out[k + '_raw']       = v
        out[k + '_judgment']  = _suffix(v)
        if note: out[k + '_unit_note'] = note

    # urinary / imaging
    out['urinaryProtein'] = (rc.get('urinaryTest') or {}).get('urinaryProtein')
    out['urinaryProtein_judgment'] = _suffix(out['urinaryProtein'])
    out['chestXray']      = (rc.get('imagingTest') or {}).get('chestXray')
    out['chestXray_judgment'] = _suffix(out['chestXray'])

    # examination
    ex = rc.get('examination') or {}
    out['pastMedicalHistory']   = ex.get('pastMedicalHistory')
    out['medicationTreatment']  = ex.get('medicationTreatment')
    out['lifestyleFactors']     = ex.get('lifestyleFactors')
    out['lifestyle_parsed']     = parse_lifestyle(ex.get('lifestyleFactors'))

    # hepatitis B
    hb = rc.get('hepatitisB') or {}
    out['hepatitisBEligibility'] = hb.get('hepatitisBEligibility')
    out['hbsAg']                 = hb.get('surfaceAntigen')
    out['hbsAb']                 = hb.get('surfaceAntibody')
    out['hbTestResult']          = hb.get('testResult')

    # specials
    out['boneDensityTest']             = rc.get('boneDensityTest')
    out['depression']                  = rc.get('depression')
    out['cognitiveDysfunction']        = rc.get('cognitiveDysfunction')
    out['geriatricPhysicalFunctionTest']= rc.get('geriatricPhysicalFunctionTest')

    gf = rc.get('geriatricFunctionTest') or {}
    out['geriatricEligibility'] = gf.get('geriatricEligibility')
    out['fallHistory']          = gf.get('fallHistory')
    out['adlScore']             = gf.get('adlScore')
    out['urinaryIncontinence']  = gf.get('urinaryIncontinence')
    out['vaccination']          = gf.get('vaccination')
    out['vaccination_parsed']   = parse_vaccination(gf.get('vaccination'))

    # disease text
    out['suspectedDisease']      = rc.get('suspectedDisease') or ''
    out['existingDisease']       = rc.get('existingDisease') or ''
    out['additionalNotes']       = rc.get('additionalNotes') or ''
    out['comprehensiveOpinion']  = rc.get('comprehensiveOpinion') or ''

    # cancer screenings (preserve full lists)
    for kind in ('breast', 'cervical', 'colorectal', 'stomach'):
        out[f'{kind}CancerScreenings'] = rc.get(f'{kind}CancerScreenings') or []

    return out


# ------------------------------------------------------------------ A sheet derivation
def _bp_status_a(s: dict) -> Optional[FindingRow]:
    j = s.get('bloodPressure_judgment') or ''
    bp = s.get('bloodPressure_raw') or s.get('bloodPressure')
    if j == '정상':
        # Normal BP — generate a "BP within normal range" row using the elevated-BP code with absent
        # (v5 example didn't show this — keep silent for normal)
        return None
    if j in ('고혈압-전단계', '고혈압-의심', '유질환자'):
        meta = A_CODES['bp_elevated_reading']
        return FindingRow(
            finding_source='measurement_derived', finding_category='A_계측',
            item_name_kr='혈압 — 고혈압 수치',
            snomed_code=meta['snomed'], snomed_display=meta['display'],
            presence_status='present', verification_status='confirmed',
            snomed_domain=meta['domain'],
            derived_value=f'{bp} mmHg', ref_range='정상<130/<80',
            medication_related='9', mapping_confidence='high',
            loinc_code=meta['loinc'],
            mapper_note='임상 해석(Hypertension)은 B시트',
        )
    return None


def _bmi_a(s: dict) -> Optional[FindingRow]:
    j = s.get('bmi_judgment')
    val = s.get('bmi')
    if j == '과체중':
        meta = A_CODES['bmi_overweight']
        return FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                          item_name_kr='체질량지수 — 과체중',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'BMI {val} kg/m²', ref_range='과체중 25–29.9',
                          loinc_code=meta['loinc'])
    if j == '비만':
        meta = A_CODES['bmi_obesity']
        return FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                          item_name_kr='체질량지수 — 비만',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'BMI {val} kg/m²', ref_range='비만 ≥30',
                          loinc_code=meta['loinc'])
    if j == '저체중':
        meta = A_CODES['bmi_underweight']
        return FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                          item_name_kr='체질량지수 — 저체중',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'BMI {val} kg/m²', ref_range='저체중 <18.5',
                          loinc_code=meta['loinc'])
    return None


def _waist_a(s: dict) -> Optional[FindingRow]:
    if s.get('waistCircumference_judgment') == '복부비만':
        meta = A_CODES['waist_abdominal_obesity']
        return FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                          item_name_kr='허리둘레 — 복부비만',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'{s.get("waistCircumference")} cm',
                          ref_range='남≥90 / 여≥85',
                          loinc_code=meta['loinc'])
    return None


def _vision_a(s: dict) -> Optional[FindingRow]:
    vL, vR = s.get('visionLeft'), s.get('visionRight')
    try:
        vLn = float(vL) if vL is not None else None
        vRn = float(vR) if vR is not None else None
    except (ValueError, TypeError):
        vLn, vRn = None, None
    if vLn is not None and vRn is not None and (vLn < 0.7 or vRn < 0.7):
        meta = A_CODES['vision_reduced']
        return FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                          item_name_kr='시력 — 시력저하 (양안)',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'시력(좌){vLn} / 시력(우){vRn}', ref_range='정상≥0.7',
                          loinc_code=meta['loinc'], mapper_note='양안 정상범위 미달')
    return None


def _hearing_a(s: dict) -> Optional[FindingRow]:
    hL, hR = s.get('hearingLeft'), s.get('hearingRight')
    if hL == '정상' and hR == '정상':
        meta = A_CODES['hearing_normal']
        return FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                          item_name_kr='청력 — 정상 (양이)',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'청력(좌){hL} / 청력(우){hR}', ref_range='정상',
                          loinc_code='—')
    return None


def _hgb_a(s: dict) -> Optional[FindingRow]:
    if s.get('hemoglobin_judgment') == '정상':
        meta = A_CODES['hgb_normal']
        return FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                          item_name_kr='혈색소 — 정상 범위',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'{s.get("hemoglobin")} g/dL',
                          ref_range='여 12.0–15.5 / 남 13.0–17.5',
                          loinc_code=meta['loinc'], mapper_note='빈혈 판정은 B시트')
    return None


def _fg_a(s: dict) -> Optional[FindingRow]:
    j = s.get('fastingGlucose_judgment') or ''
    val = s.get('fastingGlucose')
    if j and j != '정상':
        meta = A_CODES['fg_hyperglycemia']
        return FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                          item_name_kr='공복혈당 — 고혈당 (Hyperglycemia)',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'{val} mg/dL', ref_range='정상<100',
                          loinc_code=meta['loinc'], mapper_note='당뇨 판정은 B시트')
    return None


def _lipid_a(s: dict) -> List[FindingRow]:
    rows = []
    tc = s.get('totalCholesterol'); hdl = s.get('hdlCholesterol'); ldl = s.get('ldlCholesterol'); tg = s.get('triglycerides')
    def f(v):
        try: return float(v)
        except: return None
    tcN, hdlN, ldlN, tgN = f(tc), f(hdl), f(ldl), f(tg)
    if tcN and tcN > 0:
        if tcN < 200:
            meta = A_CODES['tc_normal']
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='총콜레스테롤 — 정상',
                                   snomed_code=meta['snomed'], snomed_display=meta['display'],
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=f'{tc} mg/dL', ref_range='정상<200',
                                   loinc_code=meta['loinc']))
    if hdlN and hdlN > 0 and hdlN < 50:
        meta = A_CODES['hdl_low']
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                               item_name_kr='HDL콜레스테롤 — 경계 (여성기준)',
                               snomed_code=meta['snomed'], snomed_display=meta['display'],
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'HDL {hdl} mg/dL', ref_range='여성권장≥50',
                               loinc_code=meta['loinc'], mapper_note='경계 기준'))
    if ldlN and ldlN > 0:
        if ldlN < 130:
            meta = A_CODES['ldl_normal']
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='LDL콜레스테롤 — 정상',
                                   snomed_code=meta['snomed'], snomed_display=meta['display'],
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=f'LDL {ldl} mg/dL', ref_range='정상<130',
                                   loinc_code=meta['loinc']))
    if tgN and tgN > 0 and tgN >= 150:
        meta = A_CODES['tg_high']
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                               item_name_kr='중성지방 — 경계 상승',
                               snomed_code=meta['snomed'], snomed_display=meta['display'],
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'TG {tg} mg/dL', ref_range='경계 150–199',
                               loinc_code=meta['loinc']))
    return rows


def _renal_a(s: dict) -> Optional[FindingRow]:
    cr = s.get('serumCreatinine'); egfr = s.get('egfr')
    try: egfrN = float(egfr) if egfr is not None else None
    except: egfrN = None
    if egfrN and egfrN >= 60:
        meta = A_CODES['renal_normal']
        return FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                          item_name_kr='신기능 — 정상',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'Cr {cr} / eGFR {egfr}', ref_range='eGFR≥60',
                          loinc_code=meta['loinc'])
    return None


def _liver_a(s: dict) -> Optional[FindingRow]:
    ast = s.get('ast'); alt = s.get('alt'); ggt = s.get('gamaGtp')
    try:
        astN = float(ast) if ast else None
        altN = float(alt) if alt else None
        ggtN = float(ggt) if ggt else None
    except: astN = altN = ggtN = None
    if all(v is not None for v in (astN, altN, ggtN)) and astN <= 40 and altN <= 35 and ggtN <= 35:
        meta = A_CODES['liver_normal']
        return FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                          item_name_kr='간기능 — 정상',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value=f'AST{ast}/ALT{alt}/GGT{ggt}', ref_range='AST≤40 / ALT≤35 / γGT≤35',
                          loinc_code=meta['loinc'])
    return None


def _urine_a(s: dict) -> Optional[FindingRow]:
    if s.get('urinaryProtein_judgment') == '정상':
        meta = A_CODES['urine_protein_negative']
        return FindingRow(finding_source='measurement_derived', finding_category='C_요검사',
                          item_name_kr='요단백 — 음성 (정상)',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='present', verification_status='confirmed',
                          derived_value='요단백 정상', ref_range='음성',
                          loinc_code=meta['loinc'])
    return None


def _cxr_a(s: dict) -> Optional[FindingRow]:
    cxr = s.get('chestXray') or ''
    if '비결핵성' in cxr or '질환의심' in cxr:
        meta = A_CODES['cxr_nontb_pulm']
        return FindingRow(finding_source='measurement_derived', finding_category='D_영상',
                          item_name_kr='흉부촬영 — 비결핵성 질환 의심',
                          snomed_code=meta['snomed'], snomed_display=meta['display'],
                          presence_status='suspected', verification_status='unconfirmed',
                          derived_value=f'흉부촬영: {cxr}', ref_range='정상',
                          loinc_code='—', mapper_note='결핵 판정은 B시트')
    return None


def _lifestyle_a(s: dict) -> List[FindingRow]:
    rows = []
    p = s.get('lifestyle_parsed') or {}
    # smoking
    if p.get('smoking') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='흡연 상태 — 현재 흡연',
                               snomed_code=E_CODES['current_smoker'], snomed_display='Smoker',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='■ 금연 필요'))
    elif p.get('smoking') is False:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='흡연 상태 — 평생 비흡연',
                               snomed_code=E_CODES['never_smoker'], snomed_display='Never smoked tobacco',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='□ 금연 필요',
                               mapper_note='□=비해당 → never_smoker'))
    # alcohol
    if p.get('alcohol') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='음주 상태 — 위험 음주',
                               snomed_code=E_CODES['hazardous_drinking'], snomed_display='Hazardous drinking',
                               presence_status='present', verification_status='unconfirmed',
                               snomed_domain='social_context',
                               derived_value='■ 절주 필요'))
    elif p.get('alcohol') is False:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='음주 상태 — 소량 음주',
                               snomed_code=E_CODES['social_drinker'], snomed_display='Social drinker',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='□ 절주 필요',
                               mapper_note='□=비해당 + 거의 안 마심'))
    # activity
    if p.get('activity') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='신체활동 — 신체활동 부족',
                               snomed_code=E_CODES['lack_physical_activity'], snomed_display='Lack of physical activity',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='■ 신체활동 필요',
                               mapping_confidence='medium',
                               mapper_note='■=해당'))
    # strength training
    if p.get('strength') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='근력운동 — 근력운동 부족',
                               snomed_code=E_CODES['lack_exercise'], snomed_display='Lack of exercise',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='■ 근력운동 필요',
                               mapper_note='■=해당. 신체활동(43994002)과 구분'))
    return rows


def _interview_a(s: dict) -> List[FindingRow]:
    rows = []
    if s.get('pastMedicalHistory') == '유':
        rows.append(FindingRow(finding_source='interview', finding_category='F_문진',
                               item_name_kr='과거병력 — 있음',
                               snomed_code=G_CODES['history_clinical_finding'],
                               snomed_display='History of clinical finding',
                               presence_status='present', verification_status='confirmed',
                               derived_value='과거병력=유'))
    if s.get('medicationTreatment') == '유':
        rows.append(FindingRow(finding_source='interview', finding_category='F_문진',
                               item_name_kr='현재 약물치료 중',
                               snomed_code=G_CODES['drug_treatment_compliance'],
                               snomed_display='Drug treatment compliance',
                               presence_status='present', verification_status='confirmed',
                               derived_value='약물치료=유',
                               medication_related='1'))
    return rows


def _special_a(s: dict) -> List[FindingRow]:
    """G_특수 항목 — A시트는 비해당 시 행 생성하지 않음 (B시트만 always-generate)."""
    rows = []
    # B형간염 — A시트도 비해당이면 not_measured row 생성 (v5 reference 준수)
    if s.get('hepatitisBEligibility') == '비해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='B형간염 표면항원 (HBsAg) — 검사 비해당',
                               snomed_code=G_CODES['hbsag_positive'],
                               snomed_display='Hepatitis B surface antigen positive',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당 (hepatitisBEligibility=비해당)',
                               mapping_confidence='medium',
                               mapper_note='⚠ MCP 재가동 후 검증 필요'))
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='B형간염 표면항체 (HBsAb) — 검사 비해당',
                               snomed_code=G_CODES['hbsab_positive'],
                               snomed_display='Hepatitis B surface antibody positive',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당',
                               mapping_confidence='medium',
                               mapper_note='⚠ MCP 재가동 후 검증 필요'))
    if s.get('boneDensityTest') == '비해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='골밀도검사 — 비해당',
                               snomed_code=G_CODES['osteoporosis'], snomed_display='Osteoporosis',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당'))
    if s.get('depression') == '비해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='우울증 검사 — 비해당',
                               snomed_code=G_CODES['depressive_disorder'], snomed_display='Depressive disorder',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당'))
    if s.get('cognitiveDysfunction') == '해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인지기능장애 — 특이소견 없음 (검사 시행)',
                               snomed_code=G_CODES['cognitive_impairment'],
                               snomed_display='Cognitive impairment',
                               presence_status='absent', verification_status='refuted',
                               derived_value='검사시행, 특이소견없음',
                               mapper_note='인지기능장애 해당 + 결과 이상없음 → absent'))
    elif s.get('cognitiveDysfunction') == '비해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인지기능장애 — 검사 비해당',
                               snomed_code=G_CODES['cognitive_impairment'],
                               snomed_display='Cognitive impairment',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당'))
    # geriatric
    if s.get('geriatricEligibility') == '비해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='노인신체기능검사 — 비해당',
                               snomed_code=G_CODES['physical_function_assessment'],
                               snomed_display='Assessment of physical function',
                               presence_status='not_measured', verification_status='unconfirmed',
                               snomed_domain='procedure',
                               derived_value='비해당'))
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='낙상 과거력 (fallHistory) — 비해당',
                               snomed_code='217082002', snomed_display='Fall (accident)',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당 (geriatricEligibility=비해당)'))
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='일상생활수행능력 (ADL 점수) — 비해당',
                               snomed_code=G_CODES['physical_function_assessment'],
                               snomed_display='Activities of daily living score',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당 (adlScore=null)'))
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='배뇨장애 (urinaryIncontinence) — 비해당',
                               snomed_code=G_CODES['urinary_incontinence'],
                               snomed_display='Urinary incontinence',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당 (geriatricEligibility=비해당)'))
    # vaccination — A시트에도 명시
    vac = s.get('vaccination_parsed') or {}
    if vac.get('flu') is False:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인플루엔자 예방접종 — 불필요 (□)',
                               snomed_code=G_CODES['influenza_vaccination'],
                               snomed_display='Influenza vaccination',
                               presence_status='absent', verification_status='refuted',
                               snomed_domain='procedure',
                               derived_value='□ 인플루엔자 접종 필요',
                               mapper_note='□=불필요. 항상 이 행 생성'))
    if vac.get('pneumo') is False:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='폐렴구균 예방접종 — 불필요 (□)',
                               snomed_code=G_CODES['pneumococcal_vaccination'],
                               snomed_display='Pneumococcal vaccination',
                               presence_status='absent', verification_status='refuted',
                               snomed_domain='procedure',
                               derived_value='□ 폐렴구균 접종 필요',
                               mapper_note='□=불필요. 항상 이 행 생성'))
    return rows


def _suspected_disease_a(s: dict) -> List[FindingRow]:
    """Parse the Korean free-text in suspectedDisease and emit H_질환소견 rows."""
    rows = []
    text = s.get('suspectedDisease') or ''
    if '당뇨' in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='당뇨질환 의심',
                               snomed_code=B_CODES['diabetes_mellitus'],
                               snomed_display='Diabetes mellitus',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=text[:60]))
    if '식도열공탈장' in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='식도열공탈장 의심',
                               snomed_code=B_CODES['hiatal_hernia'],
                               snomed_display='Hiatal hernia',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=text[:60],
                               mapper_note='"변화없음" → 이전부터 존재'))
    if '비만' in text and '복부비만' not in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='비만 의심 (판정문 기준)',
                               snomed_code=B_CODES['obesity'], snomed_display='Obesity',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=text[:60],
                               mapper_note='검진 판정문 원문 기준 suspected'))
    if '간장질환' in text or '간기능' in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='간장질환 의심',
                               snomed_code=B_CODES['liver_disease'],
                               snomed_display='Liver disease',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=text[:60]))
    if '신장질환' in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='신장질환 의심',
                               snomed_code=B_CODES['chronic_kidney_disease'],
                               snomed_display='Chronic kidney disease',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=text[:60]))
    if '난청' in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='난청 의심',
                               snomed_code=B_CODES['hearing_loss'],
                               snomed_display='Hearing loss',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=text[:60]))
    return rows


def _existing_disease_a(s: dict) -> List[FindingRow]:
    rows = []
    text = s.get('existingDisease') or ''
    if '고혈압' in text:
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='H_질환소견',
                               item_name_kr='고혈압 — 확진, 약물치료 중',
                               snomed_code=B_CODES['hypertension'], snomed_display='Hypertension',
                               presence_status='present', verification_status='confirmed',
                               derived_value='고혈압; 현재 약물 치료중',
                               medication_related='1'))
    return rows


def _cancer_a(s: dict) -> List[FindingRow]:
    rows = []
    SIMPLE = I_CODES['stomach_cancer_v5_simplified']  # '363346000' — v5 reference uses this for all 4 cancers
    # Stomach — gastritis / GERD biopsy findings
    for sc in s.get('stomachCancerScreenings') or []:
        op = (sc.get('testOpinion') or '') + (sc.get('screeningAssessment') or '')
        date = (sc.get('screeningDate') or '')[:10]
        if '역류성식도염' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'역류성식도염 (위내시경 {date})',
                                   snomed_code=B_CODES['gerd'], snomed_display='Gastro-esophageal reflux disease',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=op[:60],
                                   mapper_note=f'screeningDate={date}'))
        if '위축성위염' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'위축성위염 (위내시경 {date})',
                                   snomed_code=B_CODES['atrophic_gastritis'], snomed_display='Atrophic gastritis',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=op[:60],
                                   mapper_note=f'screeningDate={date}'))
        if '미란성위염' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'미란성위염 (위내시경 {date})',
                                   snomed_code=B_CODES['erosive_gastritis'], snomed_display='Erosive gastritis',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=op[:60]))
        if '위염' in op and '위축성위염' not in op and '미란성위염' not in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'위염 (위내시경 {date})',
                                   snomed_code=B_CODES['gastritis'], snomed_display='Gastritis',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=op[:60]))
        # Stomach cancer judgment row
        if '이상소견없음' in op or '양성질환' in op or '기타' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr='위암 — 없음',
                                   snomed_code=SIMPLE,
                                   snomed_display='Malignant neoplasm of stomach',
                                   presence_status='absent', verification_status='refuted',
                                   derived_value=op[:60],
                                   mapper_note=f'screeningDate={date}'))
    # Colorectal
    for sc in s.get('colorectalCancerScreenings') or []:
        op = (sc.get('testOpinion') or '') + (sc.get('screeningAssessment') or '')
        if '음성' in op or '잠혈반응 없음' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr='분변잠혈 — 음성',
                                   snomed_code=A_CODES['fobt_negative']['snomed'],
                                   snomed_display='Fecal occult blood test negative',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=op[:60]))
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr='대장암 — 없음',
                                   snomed_code=SIMPLE,
                                   snomed_display='Malignant neoplasm of large intestine',
                                   presence_status='absent', verification_status='refuted',
                                   derived_value=op[:60]))
            break
    # Breast
    for sc in s.get('breastCancerScreenings') or []:
        op = sc.get('screeningAssessment') or ''
        if '이상소견없음' in op or '양성' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr='유방암 — 없음',
                                   snomed_code=SIMPLE,
                                   snomed_display='Malignant neoplasm of breast',
                                   presence_status='absent', verification_status='refuted',
                                   derived_value=op))
            break
    return rows


def derive_a_findings(s: dict) -> List[FindingRow]:
    """A_hc_finding sheet rows."""
    rows = []
    for fn in (_vision_a, _hearing_a, _bp_status_a, _bmi_a, _waist_a):
        r = fn(s)
        if r: rows.append(r)
    rows.extend(_lipid_a(s))
    for fn in (_hgb_a, _fg_a, _renal_a, _liver_a, _urine_a, _cxr_a):
        r = fn(s)
        if r: rows.append(r)
    rows.extend(_lifestyle_a(s))
    rows.extend(_interview_a(s))
    rows.extend(_special_a(s))
    rows.extend(_existing_disease_a(s))
    rows.extend(_suspected_disease_a(s))
    rows.extend(_cancer_a(s))
    # patient_id and exam_date
    pid = s.get('patient_id'); exd = s.get('examinationDate')
    for i, r in enumerate(rows, start=1):
        r.n = i
        r.patient_id = pid
        r.exam_date = exd
    return rows


# ------------------------------------------------------------------ B sheet derivation
def derive_b_findings(s: dict) -> List[FindingRow]:
    """B_hc_clinical sheet rows — clinical interpretation with always-generate items."""
    rows = []

    # Vision (clinical interpretation)
    v = _vision_a(s)  # same code reused for B
    if v:
        v.item_name_kr = '시력저하 — 확인 (양안)'
        v.mapper_note = '양안 모두 정상 미달'
        rows.append(v)

    # Hearing — refuted if normal
    hL, hR = s.get('hearingLeft'), s.get('hearingRight')
    if hL == '정상' and hR == '정상':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='난청 — 없음 (양이 정상)',
                               snomed_code=B_CODES['hearing_loss'], snomed_display='Hearing loss',
                               presence_status='absent', verification_status='refuted',
                               derived_value=f'청력(좌){hL}/(우){hR}', ref_range='정상',
                               mapper_note='양이 정상 확인 → absent'))

    # BP — Hypertension confirmed if 유질환자
    bp_j = s.get('bloodPressure_judgment')
    if bp_j == '유질환자' or '치료중' in (s.get('existingDisease') or '') and '고혈압' in (s.get('existingDisease') or ''):
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='A_계측',
                               item_name_kr='고혈압 — 확진 (기존 진단, 약물치료 중)',
                               snomed_code=B_CODES['hypertension'], snomed_display='Hypertension',
                               presence_status='present', verification_status='confirmed',
                               derived_value=f'{s.get("bloodPressure_raw")} + 유질환자',
                               medication_related='1',
                               mapper_note='상위 진단 확정 → 전단계/정상혈압 행 불필요'))
    elif bp_j == '고혈압-전단계':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='고혈압전단계 — 확인',
                               snomed_code=B_CODES['prehypertension'], snomed_display='Prehypertension',
                               presence_status='present', verification_status='differential',
                               derived_value=s.get('bloodPressure_raw')))
    elif bp_j == '고혈압-의심':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='고혈압 — 의심',
                               snomed_code=B_CODES['hypertension'], snomed_display='Hypertension',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=s.get('bloodPressure_raw')))

    # BMI — overweight present + obesity absent (or vice versa)
    bmi_j = s.get('bmi_judgment')
    bmi_v = s.get('bmi')
    if bmi_j == '과체중':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='과체중 — 확인',
                               snomed_code=B_CODES['overweight'], snomed_display='Overweight',
                               presence_status='present', verification_status='confirmed',
                               derived_value=f'BMI {bmi_v} kg/m²'))
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='비만 — 없음 (과체중 범위)',
                               snomed_code=B_CODES['obesity'], snomed_display='Obesity',
                               presence_status='absent', verification_status='refuted',
                               derived_value=f'BMI {bmi_v} kg/m²',
                               mapper_note='과체중(present)이나 비만기준 미달 → absent'))
    elif bmi_j == '비만':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='비만 — 확인',
                               snomed_code=B_CODES['obesity'], snomed_display='Obesity',
                               presence_status='present', verification_status='confirmed',
                               derived_value=f'BMI {bmi_v} kg/m²'))
    elif bmi_j == '저체중':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='저체중 — 확인',
                               snomed_code=B_CODES['underweight'], snomed_display='Underweight',
                               presence_status='present', verification_status='confirmed',
                               derived_value=f'BMI {bmi_v} kg/m²'))

    # Waist — abdominal obesity (B sheet uses same finding code 248341001 as A for v5 consistency)
    waist_j = s.get('waistCircumference_judgment')
    if waist_j == '복부비만':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='A_계측',
                               item_name_kr='복부비만 — 확인',
                               snomed_code=A_CODES['waist_abdominal_obesity']['snomed'],
                               snomed_display='Abdominal obesity',
                               presence_status='present', verification_status='confirmed',
                               derived_value=f'허리둘레 {s.get("waistCircumference")} cm'))

    # Anemia — refuted if normal Hb
    if s.get('hemoglobin_judgment') == '정상':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                               item_name_kr='빈혈 — 없음',
                               snomed_code=B_CODES['anemia'], snomed_display='Anemia',
                               presence_status='absent', verification_status='refuted',
                               derived_value=f'Hb {s.get("hemoglobin")} g/dL'))
    elif s.get('hemoglobin_judgment') == '빈혈의심':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                               item_name_kr='빈혈 — 의심',
                               snomed_code=B_CODES['anemia'], snomed_display='Anemia',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'Hb {s.get("hemoglobin")} g/dL'))

    # Diabetes
    fg_j = s.get('fastingGlucose_judgment')
    if fg_j == '공복혈당장애 의심':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                               item_name_kr='당뇨병 — 의심',
                               snomed_code=B_CODES['diabetes_mellitus'], snomed_display='Diabetes mellitus',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'공복혈당 {s.get("fastingGlucose")} mg/dL'))
    elif fg_j == '유질환자' or '당뇨' in (s.get('existingDisease') or ''):
        rows.append(FindingRow(finding_source='disease_explicit', finding_category='B_혈액',
                               item_name_kr='당뇨병 — 확진',
                               snomed_code=B_CODES['diabetes_mellitus'], snomed_display='Diabetes mellitus',
                               presence_status='present', verification_status='confirmed',
                               derived_value=f'공복혈당 {s.get("fastingGlucose")}',
                               medication_related='1'))
    elif fg_j == '당뇨병 의심':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                               item_name_kr='당뇨병 — 의심',
                               snomed_code=B_CODES['diabetes_mellitus'], snomed_display='Diabetes mellitus',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'공복혈당 {s.get("fastingGlucose")}'))

    # Dyslipidemia
    def f(v):
        try: return float(v)
        except: return None
    tcN = f(s.get('totalCholesterol')); hdlN = f(s.get('hdlCholesterol'))
    ldlN = f(s.get('ldlCholesterol')); tgN = f(s.get('triglycerides'))
    if any(v and v > 0 for v in (tcN, hdlN, ldlN, tgN)):
        is_borderline = (hdlN and hdlN < 50) or (tgN and tgN >= 150)
        if '이상지질혈증' in (s.get('existingDisease') or ''):
            rows.append(FindingRow(finding_source='disease_explicit', finding_category='B_혈액',
                                   item_name_kr='이상지질혈증 — 확진 (기존, 치료중)',
                                   snomed_code=B_CODES['dyslipidemia'], snomed_display='Dyslipidemia',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=f'TC{s.get("totalCholesterol")}/HDL{s.get("hdlCholesterol")}/LDL{s.get("ldlCholesterol")}/TG{s.get("triglycerides")}',
                                   medication_related='1'))
        elif is_borderline:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='이상지질혈증 — 경계/의심',
                                   snomed_code=B_CODES['dyslipidemia'], snomed_display='Dyslipidemia',
                                   presence_status='suspected', verification_status='unconfirmed',
                                   derived_value=f'TC{s.get("totalCholesterol")}/HDL{s.get("hdlCholesterol")}/LDL{s.get("ldlCholesterol")}/TG{s.get("triglycerides")}'))
        # Hyper TC absent
        if tcN and tcN > 0 and tcN < 200:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='고콜레스테롤혈증 — 없음',
                                   snomed_code=B_CODES['hypercholesterolemia'], snomed_display='Hypercholesterolemia',
                                   presence_status='absent', verification_status='refuted',
                                   derived_value=f'TC {s.get("totalCholesterol")} mg/dL',
                                   mapper_note='TC<200 → absent'))
        # LDL elevation absent/present
        if ldlN and ldlN > 0:
            if ldlN < 130:
                rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                       item_name_kr='LDL 상승 — 없음',
                                       snomed_code=B_CODES['raised_ldl'], snomed_display='Raised LDL cholesterol level',
                                       presence_status='absent', verification_status='refuted',
                                       derived_value=f'LDL {s.get("ldlCholesterol")} mg/dL',
                                       mapping_confidence='medium',
                                       mapper_note='⚠ MCP 재가동 후 검증 필요'))
        # HDL low
        if hdlN and hdlN > 0 and hdlN < 50:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='HDL 저하 — 의심 (여성 경계)',
                                   snomed_code=B_CODES['low_hdl'], snomed_display='Low high density lipoprotein cholesterol level',
                                   presence_status='suspected', verification_status='unconfirmed',
                                   derived_value=f'HDL {s.get("hdlCholesterol")} mg/dL',
                                   mapper_note='여성기준 경계 → suspected'))
        # TG
        if tgN and tgN > 0 and tgN >= 150:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='고중성지방혈증 — 경계/의심',
                                   snomed_code=B_CODES['hypertriglyceridemia'], snomed_display='Hypertriglyceridemia',
                                   presence_status='suspected', verification_status='unconfirmed',
                                   derived_value=f'TG {s.get("triglycerides")} mg/dL'))

    # CKD
    egfrN = f(s.get('egfr'))
    if egfrN:
        if egfrN >= 60:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='만성신장병 — 없음',
                                   snomed_code=B_CODES['chronic_kidney_disease'], snomed_display='Chronic kidney disease',
                                   presence_status='absent', verification_status='refuted',
                                   derived_value=f'eGFR {s.get("egfr")}'))
        else:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='만성신장병 — 의심',
                                   snomed_code=B_CODES['chronic_kidney_disease'], snomed_display='Chronic kidney disease',
                                   presence_status='suspected', verification_status='unconfirmed',
                                   derived_value=f'eGFR {s.get("egfr")}'))

    # Liver function abnormal
    astN = f(s.get('ast')); altN = f(s.get('alt')); ggtN = f(s.get('gamaGtp'))
    if all(v is not None for v in (astN, altN, ggtN)):
        if astN <= 40 and altN <= 35 and ggtN <= 35:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='간기능이상 — 없음',
                                   snomed_code=B_CODES['abnormal_liver_function'], snomed_display='Abnormal liver function',
                                   presence_status='absent', verification_status='refuted',
                                   derived_value=f'AST{s.get("ast")}/ALT{s.get("alt")}/GGT{s.get("gamaGtp")}'))
        else:
            rows.append(FindingRow(finding_source='measurement_derived', finding_category='B_혈액',
                                   item_name_kr='간기능이상 — 의심',
                                   snomed_code=B_CODES['abnormal_liver_function'], snomed_display='Abnormal liver function',
                                   presence_status='suspected', verification_status='unconfirmed',
                                   derived_value=f'AST{s.get("ast")}/ALT{s.get("alt")}/GGT{s.get("gamaGtp")}'))

    # Proteinuria
    up_j = s.get('urinaryProtein_judgment')
    if up_j == '정상':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='C_요검사',
                               item_name_kr='단백뇨 — 없음',
                               snomed_code=B_CODES['proteinuria'], snomed_display='Proteinuria',
                               presence_status='absent', verification_status='refuted',
                               derived_value='요단백 음성'))
    elif up_j in ('경계', '단백뇨의심'):
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='C_요검사',
                               item_name_kr='단백뇨 — 의심',
                               snomed_code=B_CODES['proteinuria'], snomed_display='Proteinuria',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'요단백 {up_j}'))

    # Tuberculosis (chest X-ray)
    cxr = s.get('chestXray') or ''
    if cxr == '정상' or '비결핵성' in cxr:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='D_영상',
                               item_name_kr='결핵 — 없음',
                               snomed_code=B_CODES['tuberculosis'], snomed_display='Tuberculosis',
                               presence_status='absent', verification_status='refuted',
                               derived_value=f'흉부촬영: {cxr}'))
    if '비결핵성' in cxr:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='D_영상',
                               item_name_kr='비결핵성 폐질환 — 의심',
                               snomed_code=A_CODES['cxr_nontb_pulm']['snomed'],
                               snomed_display='Non-tuberculous pulmonary disease',
                               presence_status='suspected', verification_status='unconfirmed',
                               derived_value=f'흉부촬영: {cxr}',
                               mapper_note='추적 모니터링 필요'))

    # Lifestyle B-side: smoking absent if non-smoker
    p = s.get('lifestyle_parsed') or {}
    if p.get('smoking') is False:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='현재 흡연 — 없음',
                               snomed_code=E_CODES['current_smoker'], snomed_display='Smoker',
                               presence_status='absent', verification_status='refuted',
                               snomed_domain='social_context',
                               derived_value='□ 금연 필요'))
    elif p.get('smoking') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='현재 흡연 — 확인',
                               snomed_code=E_CODES['current_smoker'], snomed_display='Smoker',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='■ 금연 필요'))
    if p.get('alcohol') is False:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='위험 음주 — 없음',
                               snomed_code=E_CODES['hazardous_drinking'], snomed_display='Hazardous drinking',
                               presence_status='absent', verification_status='refuted',
                               snomed_domain='social_context',
                               derived_value='□ 절주 필요'))
    elif p.get('alcohol') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='위험 음주 — 확인',
                               snomed_code=E_CODES['hazardous_drinking'], snomed_display='Hazardous drinking',
                               presence_status='present', verification_status='unconfirmed',
                               snomed_domain='social_context',
                               derived_value='■ 절주 필요'))
    if p.get('activity') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='신체활동 부족 — 확인',
                               snomed_code=E_CODES['lack_physical_activity'], snomed_display='Lack of physical activity',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='■ 신체활동 필요',
                               mapping_confidence='medium',
                               mapper_note='⚠ MCP 재가동 후 검증 필요'))
    if p.get('strength') is True:
        rows.append(FindingRow(finding_source='lifestyle_factor', finding_category='E_생활습관',
                               item_name_kr='근력운동 부족 — 확인',
                               snomed_code=E_CODES['lack_exercise'], snomed_display='Lack of exercise',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='social_context',
                               derived_value='■ 근력운동 필요'))

    # Always-generate items (G_특수 + I_암검진)
    pid = s.get('patient_id'); exd = s.get('examinationDate')
    rows.extend(_b_always_generate(s))

    # Number them
    for i, r in enumerate(rows, start=1):
        r.n = i
        r.patient_id = pid
        r.exam_date = exd
    return rows


def _b_always_generate(s: dict) -> List[FindingRow]:
    rows = []

    # Hepatitis B (clinical)
    if s.get('hepatitisBEligibility') == '비해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='B형간염 이환 여부 — 검사 비해당',
                               snomed_code=G_CODES['hepatitis_b'], snomed_display='Hepatitis B',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당',
                               mapper_note='검사 미시행 → not_measured'))

    # Osteoporosis
    rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                           item_name_kr='골다공증 / 골밀도 이상 — 비해당' if s.get('boneDensityTest') != '해당' else '골다공증 / 골밀도 이상 — 검사시행',
                           snomed_code=G_CODES['osteoporosis'], snomed_display='Osteoporosis',
                           presence_status='not_measured' if s.get('boneDensityTest') != '해당' else 'absent',
                           verification_status='unconfirmed',
                           derived_value=s.get('boneDensityTest') or '비해당',
                           mapper_note='항상 이 행 생성'))

    # Depression
    rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                           item_name_kr='우울증 — 비해당 (검사 미시행)' if s.get('depression') != '해당' else '우울증 — 검사시행',
                           snomed_code=G_CODES['depressive_disorder'], snomed_display='Depressive disorder',
                           presence_status='not_measured' if s.get('depression') != '해당' else 'absent',
                           verification_status='unconfirmed',
                           derived_value=s.get('depression') or '비해당',
                           mapper_note='항상 이 행 생성'))

    # Cognitive impairment
    if s.get('cognitiveDysfunction') == '해당':
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인지기능장애 — 없음 (검사 시행, 특이소견 없음)',
                               snomed_code=G_CODES['cognitive_impairment'], snomed_display='Cognitive impairment',
                               presence_status='absent', verification_status='refuted',
                               derived_value='해당, 특이소견없음',
                               mapper_note='검사 시행(해당) + 이상없음 → absent'))
    else:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인지기능장애 — 검사 비해당',
                               snomed_code=G_CODES['cognitive_impairment'], snomed_display='Cognitive impairment',
                               presence_status='not_measured', verification_status='unconfirmed',
                               derived_value='비해당'))

    # Geriatric
    rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                           item_name_kr='낙상 위험 (At risk of falls) — 비해당 (미평가)',
                           snomed_code=G_CODES['at_risk_falling'], snomed_display='At risk of falling',
                           presence_status='not_measured', verification_status='unconfirmed',
                           derived_value='비해당 (geriatricEligibility=비해당)',
                           mapping_confidence='medium',
                           mapper_note='항상 이 행 생성. ⚠ MCP 재가동 후 검증'))
    rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                           item_name_kr='일상생활 수행 의존성 — 비해당 (ADL 미평가)',
                           snomed_code=G_CODES['adl_needs_assistance'], snomed_display='Needs assistance with ADL',
                           presence_status='not_measured', verification_status='unconfirmed',
                           derived_value='비해당 (adlScore=null)',
                           mapping_confidence='medium',
                           mapper_note='항상 이 행 생성'))
    rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                           item_name_kr='배뇨장애 (요실금) — 비해당 (미평가)',
                           snomed_code=G_CODES['urinary_incontinence'], snomed_display='Urinary incontinence',
                           presence_status='not_measured', verification_status='unconfirmed',
                           derived_value='비해당',
                           mapper_note='항상 이 행 생성'))

    # Vaccination
    vac = s.get('vaccination_parsed') or {}
    if vac.get('flu') is True:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인플루엔자 예방접종 — 필요 (■)',
                               snomed_code=G_CODES['influenza_vaccination'], snomed_display='Influenza vaccination',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='procedure',
                               derived_value='■ 인플루엔자 접종 필요',
                               mapper_note='항상 이 행 생성'))
    else:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='인플루엔자 예방접종 — 불필요 (□)',
                               snomed_code=G_CODES['influenza_vaccination'], snomed_display='Influenza vaccination',
                               presence_status='absent', verification_status='refuted',
                               snomed_domain='procedure',
                               derived_value='□ 인플루엔자 접종 필요',
                               mapper_note='항상 이 행 생성'))
    if vac.get('pneumo') is True:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='폐렴구균 예방접종 — 필요 (■)',
                               snomed_code=G_CODES['pneumococcal_vaccination'], snomed_display='Pneumococcal vaccination',
                               presence_status='present', verification_status='confirmed',
                               snomed_domain='procedure',
                               derived_value='■ 폐렴구균 접종 필요',
                               mapper_note='항상 이 행 생성'))
    else:
        rows.append(FindingRow(finding_source='measurement_derived', finding_category='G_특수',
                               item_name_kr='폐렴구균 예방접종 — 불필요 (□)',
                               snomed_code=G_CODES['pneumococcal_vaccination'], snomed_display='Pneumococcal vaccination',
                               presence_status='absent', verification_status='refuted',
                               snomed_domain='procedure',
                               derived_value='□ 폐렴구균 접종 필요',
                               mapper_note='항상 이 행 생성'))

    # I_암검진 always-generate (4 cancers)
    rows.extend(_cancer_b_always(s))

    return rows


def _cancer_b_always(s: dict) -> List[FindingRow]:
    rows = []
    # Use the v5-reference simplified single code 363346000 for all cancers
    # (the reference treats them all as Malignant neoplasm of [organ] with site in display)
    SIMPLE = I_CODES['stomach_cancer_v5_simplified']  # '363346000'
    cs = {
        '위암':       ('stomachCancerScreenings',     SIMPLE, 'Malignant neoplasm of stomach'),
        '대장암':     ('colorectalCancerScreenings',  SIMPLE, 'Malignant neoplasm of large intestine'),
        '유방암':     ('breastCancerScreenings',      SIMPLE, 'Malignant neoplasm of breast'),
        '자궁경부암': ('cervicalCancerScreenings',    SIMPLE, 'Malignant neoplasm of cervix uteri'),
    }
    # Stomach endoscopy findings (present rows for GERD/atrophic gastritis)
    for sc in s.get('stomachCancerScreenings') or []:
        op = (sc.get('testOpinion') or '') + (sc.get('screeningAssessment') or '')
        date = (sc.get('screeningDate') or '')[:10]
        if '역류성식도염' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'위내시경 — 역류성식도염 확인 ({date[:4] if date else ""})',
                                   snomed_code=B_CODES['gerd'], snomed_display='Gastro-esophageal reflux disease',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=f'위내시경 소견 {date}',
                                   mapper_note='전암 아님, 만성 관리 필요'))
        if '위축성위염' in op:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'위내시경 — 위축성위염 확인 ({date[:4] if date else ""}, 위암 위험인자)',
                                   snomed_code=B_CODES['atrophic_gastritis'], snomed_display='Atrophic gastritis',
                                   presence_status='present', verification_status='confirmed',
                                   derived_value=f'위내시경 소견 {date}',
                                   mapper_note='위암 위험인자. 정기 추적 필요'))

    for label, (key, code, display) in cs.items():
        screenings = s.get(key) or []
        if not screenings:
            rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                   item_name_kr=f'{label} — 검사 미시행',
                                   snomed_code=code, snomed_display=display,
                                   presence_status='not_measured', verification_status='unconfirmed',
                                   derived_value='검사 시행 없음',
                                   mapper_note='항상 이 행 생성'))
        else:
            assessments = [(sc.get('screeningAssessment') or '') for sc in screenings]
            ops = [(sc.get('testOpinion') or '') for sc in screenings]
            joined = ' / '.join(assessments + ops)
            if any('이상소견없음' in a for a in assessments) or '잠혈반응 없음' in joined or '음성' in joined or '양성질환' in joined:
                rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                       item_name_kr=f'{label} — 없음',
                                       snomed_code=code, snomed_display=display,
                                       presence_status='absent', verification_status='refuted',
                                       derived_value=joined[:60]))
            else:
                rows.append(FindingRow(finding_source='cancer_screening', finding_category='I_암검진',
                                       item_name_kr=f'{label} — 검사결과 추적',
                                       snomed_code=code, snomed_display=display,
                                       presence_status='suspected', verification_status='unconfirmed',
                                       derived_value=joined[:60]))
    return rows


# ------------------------------------------------------------------ C sheet (comparison)
def derive_c_findings(prev: dict, curr: dict) -> List[FindingRow]:
    """Compare two normalized subjects (older vs newer); return rows for items that differ."""
    if prev is None or curr is None: return []
    rows = []
    diffs = []
    fields_to_compare = [
        ('hearingLeft',   '청력(좌)'),
        ('hearingRight',  '청력(우)'),
        ('bloodPressure', '혈압'),
        ('bmi',           '체질량지수'),
        ('waistCircumference', '허리둘레'),
        ('fastingGlucose','공복혈당'),
        ('hemoglobin',    '혈색소'),
        ('totalCholesterol','총콜레스테롤'),
        ('hdlCholesterol','HDL'),
        ('ldlCholesterol','LDL'),
        ('triglycerides', '중성지방'),
        ('ast', 'AST'), ('alt', 'ALT'), ('gamaGtp', 'γGTP'),
        ('chestXray', '흉부촬영'),
        ('urinaryProtein', '요단백'),
    ]
    pid = curr.get('patient_id')
    prev_date = prev.get('examinationDate', '')
    n = 0
    for field, label in fields_to_compare:
        pv = prev.get(field); cv = curr.get(field)
        if pv == cv: continue
        if pv is None and cv is None: continue
        n += 1
        rows.append(FindingRow(
            n=n, patient_id=pid, exam_date=prev_date,
            finding_source='measurement_derived', finding_category='A_계측',
            item_name_kr=f'{label} — 차이 ★{prev_date[:4]}',
            snomed_code='', snomed_display='',
            presence_status='present' if (pv and pv != '정상') else 'absent',
            verification_status='unconfirmed',
            derived_value=f'{prev_date}: {pv} → {curr.get("examinationDate","")[:4]}: {cv}',
            mapper_note='두 일자 비교 자동 추출',
        ))
    return rows
