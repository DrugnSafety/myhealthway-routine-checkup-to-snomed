#!/usr/bin/env python3
"""
SNOMED CT Mapper MCP Server (FastMCP)

Exposes the v5 conversion pipeline as MCP tools that can be invoked from any
Claude session without bash access. Re-uses the deterministic Python core in
`skills/snomed-ct-mapper/scripts/`.

Tools exposed:
  - convert_subject     : single-patient JSON/dict → v5 workbook
  - batch_convert       : multi-subject JSON/xlsx → directory of v5 workbooks
  - verify_workbook_v5  : schema validation + optional golden compare
  - get_codes           : read curated SNOMED+LOINC mapping table
  - normalize_subject   : flatten a routineCheckupDataResponse for inspection

Stdio transport. Configured via .mcp.json at plugin root.
"""
from __future__ import annotations
import json
import os
import sys
from typing import Optional

# Resolve plugin root (two dirs up from this file: mcp/server.py)
HERE = os.path.dirname(os.path.abspath(__file__))
PLUGIN_ROOT = os.path.normpath(os.path.join(HERE, '..'))
SCRIPTS = os.path.join(PLUGIN_ROOT, 'skills', 'snomed-ct-mapper', 'scripts')
sys.path.insert(0, SCRIPTS)

# Pull in the deterministic core
from v5_core import (                          # noqa: E402
    normalize_subject as _normalize,
    derive_a_findings, derive_b_findings, derive_c_findings,
    load_codes,
)
from v5_writer import write_workbook           # noqa: E402
from verify_workbook import verify as _verify  # noqa: E402

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    print(
        'ERROR: fastmcp not installed. Install via: pip install "mcp[cli]"',
        file=sys.stderr,
    )
    sys.exit(1)

mcp = FastMCP('snomed-ct-mapper')


def _extract_rc(payload, name=None):
    """Find a routineCheckupDataResponse dict inside any of the supported shapes."""
    if isinstance(payload, dict):
        if 'routineCheckupDataResponse' in payload:
            return payload['routineCheckupDataResponse']
        if any(k in payload for k in ('patientName', 'measurementTest', 'examination')):
            return payload
        return None
    if isinstance(payload, list):
        for entry in payload:
            rc = entry.get('routineCheckupDataResponse') if isinstance(entry, dict) else None
            if not rc:
                continue
            if name is None or rc.get('patientName') == name:
                return rc
    return None


# ──────────────────────────────────────────────────────────────────────
# Tool: convert_subject
# ──────────────────────────────────────────────────────────────────────
@mcp.tool()
def convert_subject(
    input_json_path: str,
    output_xlsx_path: str,
    patient_name: Optional[str] = None,
    compare_with_json_path: Optional[str] = None,
) -> dict:
    """Convert one patient's Korean health-checkup JSON into a v5 SNOMED CT workbook.

    Args:
        input_json_path:        Absolute path to the JSON file (single envelope or list).
        output_xlsx_path:       Absolute path where the .xlsx will be written.
        patient_name:           Optional Korean name; required only when the JSON contains many subjects.
        compare_with_json_path: Optional second JSON for comparison mode (older exam date).

    Returns:
        dict with keys: status, output, patient, patient_id, exam_date, a_rows, b_rows, c_rows, comparison.
    """
    with open(input_json_path, encoding='utf-8') as f:
        payload = json.load(f)
    rc = _extract_rc(payload, name=patient_name)
    if rc is None:
        return {'status': 'error', 'error': f'No matching subject (name={patient_name})'}

    prev = None
    if compare_with_json_path:
        with open(compare_with_json_path, encoding='utf-8') as f:
            prev_payload = json.load(f)
        prev_rc = _extract_rc(prev_payload, name=patient_name or rc.get('patientName'))
        if prev_rc:
            prev = _normalize(prev_rc)

    curr = _normalize(rc)
    a = derive_a_findings(curr)
    b = derive_b_findings(curr)
    c = derive_c_findings(prev, curr) if prev else []

    os.makedirs(os.path.dirname(os.path.abspath(output_xlsx_path)), exist_ok=True)
    write_workbook(output_xlsx_path, curr, prev, a, b, c)

    return {
        'status':     'ok',
        'output':     output_xlsx_path,
        'patient':    curr.get('patientName'),
        'patient_id': curr.get('patient_id'),
        'exam_date':  curr.get('examinationDate'),
        'a_rows':     len(a),
        'b_rows':     len(b),
        'c_rows':     len(c),
        'comparison': bool(prev),
    }


# ──────────────────────────────────────────────────────────────────────
# Tool: batch_convert
# ──────────────────────────────────────────────────────────────────────
@mcp.tool()
def batch_convert(
    input_path: str,
    output_dir: str,
    input_kind: str = 'json',
    limit: Optional[int] = None,
) -> dict:
    """Convert every subject in an input source to its own v5 workbook.

    Args:
        input_path:  JSON list of envelopes OR flattened 132-col xlsx.
        output_dir:  Directory where per-subject xlsx files are written.
        input_kind:  'json' or 'xlsx'.
        limit:       Optional cap (process first N subjects). Useful for smoke tests.

    Returns:
        dict with summary: total, ok, failed, out_dir, results (list of per-subject reports).
    """
    sys.path.insert(0, SCRIPTS)
    from batch_convert import (                        # noqa: E402
        _from_json, _from_xlsx, _group_by_patient, _pick_compare, _safe_filename,
    )
    from verify_workbook import verify as v_verify     # noqa: E402

    os.makedirs(output_dir, exist_ok=True)
    stream = _from_json(input_path) if input_kind == 'json' else _from_xlsx(input_path)
    grouped = _group_by_patient(stream)
    items = list(grouped.items())
    if limit:
        items = items[:limit]

    results = []
    n_ok, n_fail = 0, 0
    for (name, birth), records in items:
        try:
            if len(records) >= 2:
                prev_rc, curr_rc = _pick_compare(records)
                if (prev_rc.get('examinationDate') or '') == (curr_rc.get('examinationDate') or ''):
                    prev_rc = None
            else:
                curr_rc = records[0]; prev_rc = None
            curr = _normalize(curr_rc)
            prev = _normalize(prev_rc) if prev_rc else None
            a = derive_a_findings(curr)
            b = derive_b_findings(curr)
            c = derive_c_findings(prev, curr) if prev else []
            fname = _safe_filename(name, birth)
            out_path = os.path.join(output_dir, fname)
            i = 2
            while os.path.exists(out_path):
                out_path = os.path.join(output_dir, fname.replace('.xlsx', f'_{i}.xlsx'))
                i += 1
            write_workbook(out_path, curr, prev, a, b, c)
            errs, warns, _ = v_verify(out_path)
            verify_status = 'PASS' if not errs else f'ERR x {len(errs)}'
            results.append({
                'name': name, 'birth': birth,
                'exam_date': curr.get('examinationDate'),
                'comparison': bool(prev),
                'a_rows': len(a), 'b_rows': len(b), 'c_rows': len(c),
                'verify': verify_status,
                'output': out_path,
            })
            n_ok += 1
        except Exception as e:
            n_fail += 1
            results.append({'name': name, 'birth': birth, 'error': str(e)})

    return {
        'total':   len(items),
        'ok':      n_ok,
        'failed':  n_fail,
        'out_dir': output_dir,
        'results': results,
    }


# ──────────────────────────────────────────────────────────────────────
# Tool: verify_workbook_v5
# ──────────────────────────────────────────────────────────────────────
@mcp.tool()
def verify_workbook_v5(
    workbook_path: str,
    golden_path: Optional[str] = None,
    strict: bool = False,
) -> dict:
    """Validate a v5 workbook against schema rules and optionally a golden reference.

    Args:
        workbook_path: Absolute path to the .xlsx to verify.
        golden_path:   Optional absolute path to a reference .xlsx for row-count comparison.
        strict:        If True, large row-count divergence vs golden is an error (not warning).

    Returns:
        dict with: errors, warnings, info (stats + diff_vs_golden), pass.
    """
    errs, warns, info = _verify(workbook_path, golden=golden_path, strict=strict)
    return {
        'errors':   errs,
        'warnings': warns,
        'info':     info,
        'pass':     not errs,
    }


# ──────────────────────────────────────────────────────────────────────
# Tool: get_codes
# ──────────────────────────────────────────────────────────────────────
@mcp.tool()
def get_codes(section: Optional[str] = None) -> dict:
    """Return the curated SNOMED CT + LOINC code mapping table.

    Args:
        section: One of 'A_hc_finding_codes', 'B_hc_clinical_codes', 'E_lifestyle_codes',
                 'G_special_codes', 'I_cancer_codes', 'loinc_observables',
                 'always_generate_in_B_sheet', 'verification_required_v5', 'qualifier_value'.
                 If None, returns the full table.

    Returns:
        dict containing the requested section (or the full CODES.json contents).
    """
    codes = load_codes()
    if section is None:
        return codes
    if section not in codes:
        return {'error': f'unknown section: {section}', 'available': list(codes.keys())}
    return {section: codes[section]}


# ──────────────────────────────────────────────────────────────────────
# Tool: normalize_subject
# ──────────────────────────────────────────────────────────────────────
@mcp.tool()
def normalize_subject(
    input_json_path: str,
    patient_name: Optional[str] = None,
) -> dict:
    """Flatten a single patient's routineCheckupDataResponse with unit conversions applied.

    Useful for inspection before running the full conversion. Does not write any files.

    Args:
        input_json_path: Absolute path to the JSON.
        patient_name:    Optional Korean name when the JSON has multiple subjects.

    Returns:
        Flattened subject dict (with unit-converted fields and parsed lifestyle/vaccination).
    """
    with open(input_json_path, encoding='utf-8') as f:
        payload = json.load(f)
    rc = _extract_rc(payload, name=patient_name)
    if rc is None:
        return {'error': f'No matching subject (name={patient_name})'}
    return _normalize(rc)


if __name__ == '__main__':
    mcp.run()
