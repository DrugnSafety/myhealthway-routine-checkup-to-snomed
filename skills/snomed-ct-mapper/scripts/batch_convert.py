#!/usr/bin/env python3
"""batch_convert — Apply convert_to_v5 to every subject in an input source.

Inputs (mutually exclusive):
  --json LIST.json    — JSON list of envelopes; one v5 file per subject
  --xlsx FLAT132.xlsx — flattened 132-col workbook; one v5 file per data row

Output:
  --out-dir DIR       — directory where the per-subject .xlsx files are written
  --index PATH.xlsx   — optional index workbook listing all generated files

Comparison-mode is automatic: if a subject appears with two distinct
examinationDate values, the older one is passed as --compare-with.

Verification: every output is checked with verify_workbook (warnings only).
"""
from __future__ import annotations
import argparse, json, os, sys, re, traceback, unicodedata
from collections import defaultdict

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, THIS_DIR)

from v5_core import normalize_subject, derive_a_findings, derive_b_findings, derive_c_findings
from v5_writer import write_workbook
from verify_workbook import verify


def _safe_filename(name, birth):
    name = unicodedata.normalize('NFC', name or 'NA')
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    return f'{name}_{birth or "unknown"}_v5.xlsx'


def _from_json(path):
    """Yield (key, rc_dict, birth) for each subject in the JSON list."""
    with open(path, encoding='utf-8') as f:
        payload = json.load(f)
    if isinstance(payload, dict):
        rc = payload.get('routineCheckupDataResponse') or payload
        yield (rc.get('patientName'), rc, rc.get('birthDate'))
        return
    for entry in payload:
        rc = entry.get('routineCheckupDataResponse') if isinstance(entry, dict) else None
        if not rc: continue
        yield (rc.get('patientName'), rc, rc.get('birthDate'))


def _from_xlsx(path):
    """Yield (patientName, rc_like_dict, birthDate) for each data row."""
    from convert_to_v5 import _load_from_xlsx
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb.active
    for r in range(2, ws.max_row+1):
        rc = _load_from_xlsx(path, r)
        yield (rc.get('patientName'), rc, rc.get('birthDate'))


def _group_by_patient(stream):
    """Collect all rc records keyed by (name, birth) so we can detect comparison pairs."""
    grouped = defaultdict(list)
    for name, rc, birth in stream:
        if not name: continue
        grouped[(name, birth)].append(rc)
    return grouped


def _pick_compare(records):
    """Given >=2 records for the same subject, pick (older, newer) by examinationDate."""
    sorted_recs = sorted(records, key=lambda r: r.get('examinationDate') or '')
    return sorted_recs[0], sorted_recs[-1]


def _make_index(out_dir, results, index_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    THIN = Side(border_style='thin', color='B0B0B0')
    BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    HDR_FILL = PatternFill('solid', fgColor='305496')
    HDR_FONT = Font(bold=True, color='FFFFFF', name='맑은 고딕', size=10)
    wb = Workbook(); ws = wb.active; ws.title = '대상자_v5_색인'
    headers = ['#', '환자명', '생년월일', '검진일', '비교모드', 'A행수', 'B행수', 'C행수', '검증', '파일명']
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = HDR_FILL; c.font = HDR_FONT
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER
    for i, r in enumerate(results, start=1):
        row = [i, r['name'], r['birth'], r['exam_date'], '예' if r['comparison'] else '',
               r['a'], r['b'], r['c'], r['verify'], r['filename']]
        for ci, v in enumerate(row, start=1):
            cell = ws.cell(row=i+1, column=ci, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical='top')
    widths = [4, 12, 13, 13, 8, 7, 7, 7, 8, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = 'A2'
    wb.save(index_path)


def main():
    ap = argparse.ArgumentParser()
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument('--json', help='JSON list of envelopes')
    src.add_argument('--xlsx', help='Flattened 132-col xlsx')
    ap.add_argument('--out-dir', required=True, help='directory for per-subject v5 files')
    ap.add_argument('--index', help='write an index workbook')
    ap.add_argument('--limit', type=int, help='only process first N subjects (for testing)')
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    stream = _from_json(args.json) if args.json else _from_xlsx(args.xlsx)
    grouped = _group_by_patient(stream)
    items = list(grouped.items())
    if args.limit:
        items = items[:args.limit]

    results, n_ok, n_fail = [], 0, 0
    for (name, birth), records in items:
        try:
            if len(records) >= 2:
                prev_rc, curr_rc = _pick_compare(records)
                # if both have the same date, treat as single
                if (prev_rc.get('examinationDate') or '') == (curr_rc.get('examinationDate') or ''):
                    prev_rc = None
            else:
                curr_rc = records[0]; prev_rc = None
            curr = normalize_subject(curr_rc)
            prev = normalize_subject(prev_rc) if prev_rc else None
            a = derive_a_findings(curr)
            b = derive_b_findings(curr)
            c = derive_c_findings(prev, curr) if prev else []
            fname = _safe_filename(name, birth)
            out_path = os.path.join(args.out_dir, fname)
            # avoid overwrite collisions
            i = 2
            while os.path.exists(out_path):
                out_path = os.path.join(args.out_dir, fname.replace('.xlsx', f'_{i}.xlsx'))
                i += 1
            write_workbook(out_path, curr, prev, a, b, c)
            errs, warns, _ = verify(out_path)
            verify_status = 'PASS' if not errs else f'ERR×{len(errs)}'
            results.append({
                'name': name, 'birth': birth,
                'exam_date': curr.get('examinationDate'),
                'comparison': bool(prev),
                'a': len(a), 'b': len(b), 'c': len(c),
                'verify': verify_status,
                'filename': os.path.basename(out_path),
            })
            n_ok += 1
            print(f'[{n_ok:>3}/{len(items)}] OK   {name} ({birth}) → {os.path.basename(out_path)} '
                  f'(A{len(a)}/B{len(b)}/C{len(c)}) {verify_status}')
        except Exception as e:
            n_fail += 1
            print(f'[FAIL] {name} ({birth}): {e}', file=sys.stderr)
            traceback.print_exc()

    if args.index:
        _make_index(args.out_dir, results, args.index)

    summary = {'total': len(items), 'ok': n_ok, 'failed': n_fail, 'out_dir': args.out_dir}
    print('\n=== batch_convert summary ===')
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if n_fail == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
